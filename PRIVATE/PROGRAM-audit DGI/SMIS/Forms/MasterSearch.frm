VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMIS_Mis_SearchMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search For Prospect/Customer"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
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
   Icon            =   "MasterSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9210
   StartUpPosition =   1  'CenterOwner
   Begin XtremeReportControl.ReportControl lvGrid 
      Height          =   4515
      Left            =   30
      TabIndex        =   27
      Top             =   1020
      Width           =   9150
      _Version        =   655364
      _ExtentX        =   16140
      _ExtentY        =   7964
      _StockProps     =   64
      BorderStyle     =   4
      AllowColumnRemove=   0   'False
   End
   Begin VB.PictureBox picCustomer 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   60
      ScaleHeight     =   540
      ScaleWidth      =   9075
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CommandButton cmdEdit 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8430
         MouseIcon       =   "MasterSearch.frx":000C
         MousePointer    =   99  'Custom
         Picture         =   "MasterSearch.frx":015E
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   435
         Left            =   7890
         MouseIcon       =   "MasterSearch.frx":04BA
         MousePointer    =   99  'Custom
         Picture         =   "MasterSearch.frx":060C
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Add Record"
         Top             =   30
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.OptionButton optselect 
         Caption         =   "&Customer Name"
         Height          =   285
         Index           =   12
         Left            =   5355
         TabIndex        =   15
         Tag             =   "CUSTOMERNAME"
         Top             =   75
         Width           =   1680
      End
      Begin VB.OptionButton optselect 
         Caption         =   "Account &Name"
         Height          =   285
         Index           =   11
         Left            =   0
         TabIndex        =   14
         Tag             =   "ACCTNAME"
         Top             =   75
         Value           =   -1  'True
         Width           =   1515
      End
      Begin VB.OptionButton optselect 
         Caption         =   "&Email"
         Height          =   285
         Index           =   10
         Left            =   4344
         TabIndex        =   13
         Tag             =   "EMAIL"
         Top             =   75
         Width           =   825
      End
      Begin VB.OptionButton optselect 
         Caption         =   "&Telephone"
         Height          =   285
         Index           =   9
         Left            =   2916
         TabIndex        =   12
         Tag             =   "PHONE"
         Top             =   75
         Width           =   1245
      End
      Begin VB.OptionButton optselect 
         Caption         =   "A&ddress"
         Height          =   285
         Index           =   8
         Left            =   1698
         TabIndex        =   11
         Tag             =   "ADDRESS"
         Top             =   75
         Width           =   1035
      End
   End
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   60
      TabIndex        =   28
      Text            =   "XXX"
      Top             =   630
      Width           =   4245
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
      Height          =   735
      Left            =   8430
      MouseIcon       =   "MasterSearch.frx":091F
      MousePointer    =   99  'Custom
      Picture         =   "MasterSearch.frx":0A71
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Cancel"
      Top             =   5610
      Width           =   705
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
      Height          =   735
      Left            =   7740
      MouseIcon       =   "MasterSearch.frx":0DAF
      MousePointer    =   99  'Custom
      Picture         =   "MasterSearch.frx":0F01
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Select"
      Top             =   5610
      Width           =   705
   End
   Begin VB.PictureBox picMRR 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   9075
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   9075
      Begin VB.OptionButton optselect 
         Caption         =   "&Model"
         Height          =   285
         Index           =   22
         Left            =   105
         TabIndex        =   34
         Tag             =   "Model"
         Top             =   90
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.OptionButton optselect 
         Caption         =   "ModelYear"
         Height          =   285
         Index           =   21
         Left            =   2910
         TabIndex        =   33
         Tag             =   "YEER"
         Top             =   90
         Width           =   1215
      End
      Begin VB.OptionButton optselect 
         Caption         =   "&Description"
         Height          =   285
         Index           =   20
         Left            =   1365
         TabIndex        =   32
         Tag             =   "DESCRIPT"
         Top             =   90
         Width           =   1245
      End
   End
   Begin VB.PictureBox picPO 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   60
      ScaleHeight     =   540
      ScaleWidth      =   9075
      TabIndex        =   23
      Top             =   30
      Visible         =   0   'False
      Width           =   9075
      Begin VB.OptionButton optselect 
         Caption         =   "&Model"
         Height          =   285
         Index           =   19
         Left            =   1545
         TabIndex        =   26
         Tag             =   "SMIS_PO.Model"
         Top             =   90
         Width           =   1245
      End
      Begin VB.OptionButton optselect 
         Caption         =   "Source"
         Height          =   285
         Index           =   17
         Left            =   2790
         TabIndex        =   25
         Tag             =   "SMIS_PO.Source"
         Top             =   90
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optselect 
         Caption         =   "&PO NO"
         Height          =   285
         Index           =   18
         Left            =   105
         TabIndex        =   24
         Tag             =   "SMIS_PO.PO_NO"
         Top             =   90
         Width           =   1245
      End
   End
   Begin VB.PictureBox picProspect 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   60
      ScaleHeight     =   540
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9075
      Begin VB.OptionButton optselect 
         Caption         =   "&Mobile"
         Height          =   285
         Index           =   16
         Left            =   7410
         TabIndex        =   20
         Tag             =   "Mobile"
         Top             =   60
         Width           =   1095
      End
      Begin VB.OptionButton optselect 
         Caption         =   "&Prospect Name"
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Tag             =   "AcctName"
         Top             =   60
         Value           =   -1  'True
         Width           =   1680
      End
      Begin VB.OptionButton optselect 
         Caption         =   "&Email"
         Height          =   285
         Index           =   2
         Left            =   5565
         TabIndex        =   7
         Tag             =   "EMAIL"
         Top             =   60
         Width           =   825
      End
      Begin VB.OptionButton optselect 
         Caption         =   "&Telephone"
         Height          =   285
         Index           =   3
         Left            =   1815
         TabIndex        =   6
         Tag             =   "TelePhone"
         Top             =   60
         Width           =   1245
      End
      Begin VB.OptionButton optselect 
         Caption         =   "A&ddress"
         Height          =   285
         Index           =   4
         Left            =   4350
         TabIndex        =   5
         Tag             =   "ADDRESS"
         Top             =   60
         Width           =   1035
      End
      Begin VB.OptionButton optselect 
         Caption         =   "&Model"
         Height          =   285
         Index           =   0
         Left            =   3255
         TabIndex        =   4
         Tag             =   "Variant"
         Top             =   60
         Width           =   915
      End
      Begin VB.OptionButton optselect 
         Caption         =   "&SAE"
         Height          =   285
         Index           =   5
         Left            =   6525
         TabIndex        =   3
         Tag             =   "SAE"
         Top             =   60
         Width           =   825
      End
   End
   Begin VB.PictureBox picApplicant 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   60
      ScaleHeight     =   540
      ScaleWidth      =   9075
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   9075
      Begin VB.OptionButton optselect 
         Caption         =   "&Applicant Name"
         Height          =   285
         Index           =   7
         Left            =   90
         TabIndex        =   10
         Tag             =   "ApplicantName"
         Top             =   90
         Value           =   -1  'True
         Width           =   1680
      End
      Begin VB.OptionButton optselect 
         Caption         =   "&Sales Account Executive"
         Height          =   285
         Index           =   6
         Left            =   3870
         TabIndex        =   9
         Tag             =   "SAE"
         Top             =   90
         Width           =   2565
      End
   End
   Begin VB.PictureBox picQuotation 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   60
      ScaleHeight     =   540
      ScaleWidth      =   9075
      TabIndex        =   16
      Top             =   30
      Visible         =   0   'False
      Width           =   9075
      Begin VB.OptionButton optselect 
         Caption         =   "&Prospect Name"
         Height          =   285
         Index           =   15
         Left            =   1740
         TabIndex        =   19
         Tag             =   "CP.AcctName"
         Top             =   120
         Value           =   -1  'True
         Width           =   1635
      End
      Begin VB.OptionButton optselect 
         Caption         =   "&Date"
         Height          =   285
         Index           =   14
         Left            =   3720
         TabIndex        =   18
         Tag             =   "CQ.QuotationDate"
         Top             =   120
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.OptionButton optselect 
         Caption         =   "&Model Name"
         Height          =   285
         Index           =   13
         Left            =   135
         TabIndex        =   17
         Tag             =   "CQ.ModelDescript"
         Top             =   120
         Width           =   1680
      End
   End
   Begin VB.PictureBox picRelease 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   9075
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   9075
      Begin VB.OptionButton optselect 
         Caption         =   "&Customer Name"
         Height          =   285
         Index           =   25
         Left            =   1425
         TabIndex        =   38
         Tag             =   "custname"
         Top             =   90
         Width           =   1845
      End
      Begin VB.OptionButton optselect 
         Caption         =   "CS #"
         Height          =   285
         Index           =   24
         Left            =   3390
         TabIndex        =   37
         Tag             =   "IGNKEY"
         Top             =   90
         Width           =   1245
      End
      Begin VB.OptionButton optselect 
         Caption         =   "&Invoice No"
         Height          =   285
         Index           =   23
         Left            =   105
         TabIndex        =   36
         Tag             =   "SMIS_SALESORDER.VI_NO"
         Top             =   90
         Value           =   -1  'True
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmSMIS_Mis_SearchMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents CusForm                                            As frmAllCustomer
Attribute CusForm.VB_VarHelpID = -1
Public Event SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
Public Event NoSelectionMade()
Dim strCondition                                                      As String
Dim SEARCHFOR                                                         As String

Function GetControlTag(PicBox As PictureBox) As String
    Dim i                                                             As Integer
    For i = 0 To optselect.Count - 1

        If optselect(i).Value = True And optselect(i).Container.Name = PicBox.Name Then
            GetControlTag = optselect(i).Tag

            Exit Function
        End If
    Next

End Function

Sub SearchForPO()
    SEARCHFOR = "PO"
End Sub

Sub SearchForApplication()
    SEARCHFOR = "APPLICATION"
End Sub

Sub SearchForMRR()
    SEARCHFOR = "MRR"
End Sub

Sub SearchForRELEASED()
    SEARCHFOR = "RELEASED"
End Sub

Sub SearchForCustomers()
    SEARCHFOR = "CUSTOMER"
End Sub

Sub SearchForProspects(Optional ByVal Conditionstring As String = vbNullString)
    SEARCHFOR = "PROSPECT"
    If Conditionstring <> vbNullString Then
        strCondition = " AND  " & Conditionstring
    End If
End Sub

Sub SearchForQuotation()
    SEARCHFOR = "QUOTE"
End Sub

Sub SearchForSalesOrder()

End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "CUSTOMER") = False Then Exit Sub
    Set CusForm = New frmAllCustomer
    Load CusForm
    CusForm.cmdAdd.Value = True
    CusForm.Show 1
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo ErrorCode:

    RaiseEvent NoSelectionMade
    Unload Me
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "CUSTOMER") = False Then Exit Sub

    If lvGrid.SelectedRows.Count <= 0 Then Exit Sub

    Set CusForm = New frmAllCustomer
    CusForm.AddEditCustomer (lvGrid.SelectedRows(0).Record(6).Value)
    CusForm.Show 1

End Sub

Private Sub cmdSelect_Click()

    On Error GoTo ErrorCode:

    If lvGrid.Records.Count = 0 Then
        MsgBox "Select Record Please"
        Exit Sub
    End If
    'LOAN APPLICATIONS
    If picApplicant.Visible = True Then
        'DateApplied, ApplicantName, UnitModel,  SAE, ApplicationType, ID, ProspectId, APlCode
        If lvGrid.SelectedRows(0).Record(4).Value = "Corporate" Then
            'CORPORATE LOAN APPLICATIONS
            RaiseEvent SelectionMade(gconDMIS.Execute("Select * , 'C' as ApplicationType from SMIS_LoanCorp WHERE ID= " & lvGrid.SelectedRows(0).Record(5).Value), "APPLICATIONCORP")
        Else
            'INDIVIDUAL LOAN APPLICATIONS
            'DateApplied, NAME, MODEL ,Ind_LoanApl_SAE, ID
            RaiseEvent SelectionMade(gconDMIS.Execute("Select * , 'I' as ApplicationType  from SMIS_LoanIndiv WHERE ID= " & lvGrid.SelectedRows(0).Record(5).Value), "APPLICATIONINDIV")
        End If
    ElseIf picCustomer.Visible = True Then
        'CUstomerName,Address, MOBILE, Email,Custype,CUSTID,CUSCDE
        'CUSTOMERS
        RaiseEvent SelectionMade(gconDMIS.Execute("Select * from ALL_CUSTOMER WHERE ID= " & lvGrid.SelectedRows(0).Record(5).Value), "CUSTOMER")
        
    ElseIf picProspect.Visible = True Then
        ' LogInitialInquiry,  AcctName ,  Variant,SAE , TelePhone ,  Mobile,Address ,CUSCDE, ProspectType, ProspectID
        'PROSPECTS
        RaiseEvent SelectionMade(gconDMIS.Execute("Select * from CRIS_PROSPECTS WHERE ProspectID= " & lvGrid.SelectedRows(0).Record(8).Value), "PROSPECT")
    ElseIf picQuotation.Visible = True Then
        'CQ.QuotationDate, CP.AcctName, CQ.ModelDescript, CP.ContactPerson,CP.Telephone , CP.Email, CQ.ID
        'QUOTATIONS
        RaiseEvent SelectionMade(gconDMIS.Execute("Select * from CRIS_Quotation WHERE LOGID= " & lvGrid.SelectedRows(0).Record(6).Value), "QUOTATION")
    ElseIf picPO.Visible = True Then
        '"DateOrdered, PONO,  Description, ModelYear, ID"
        RaiseEvent SelectionMade(gconDMIS.Execute("Select * from SMIS_PO WHERE ID= " & lvGrid.SelectedRows(0).Record(5).Value), "PO")
    ElseIf picMRR.Visible = True Then
        RaiseEvent SelectionMade(gconDMIS.Execute("Select * from SMIS_MRRINV WHERE ID= " & lvGrid.SelectedRows(0).Record(4).Value), "MRR")
    ElseIf picRelease.Visible = True Then
        RaiseEvent SelectionMade(gconDMIS.Execute("Select IGNKEY_NO as IGNKEY, VI_NO FROM SMIS_SALESORDER WHERE VI_NO= '" & lvGrid.SelectedRows(0).Record(3).Value & "'"), "MRR")

    End If
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub CusForm_ChangedData(xCUSCODE As String)
    txtsearch_Change
    Unload CusForm
    Set CusForm = Nothing
End Sub

Private Sub Form_Activate()
    txtSearch = ""
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Trim(txtSearch) <> "" Then
            On Error Resume Next
            txtSearch.SetFocus
        Else

            cmdCancel_Click

        End If
    End If

End Sub

Private Sub Form_Load()
    CenterMe Me, frmMain, 0
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    Select Case SEARCHFOR
        Case "PROSPECT"
            ReportControlAddColumnHeader lvGrid, "Date,ProspectName, Model, SAE, Phone,Mobile,Address"
            ReportControlPaintManager lvGrid
            ResizeColumnHeader lvGrid, "15,30,20,20,15,15,20"
            lvGrid.Columns(4).Alignment = xtpAlignmentCenter
            lvGrid.Columns(5).Alignment = xtpAlignmentCenter
            picProspect.Visible = True
            picProspect.Top = 0
            picProspect.Left = 90
            Me.Caption = "Search for Prospects"
        Case "APPLICATION"
            ResizeColumnHeader lvGrid, "10, 40,30,20"
            ReportControlAddColumnHeader lvGrid, "DateApplied,ApplicantName, Model, SAE"
            ReportControlPaintManager lvGrid
            picApplicant.Visible = True
            picApplicant.Top = 0
            picApplicant.Left = 90
            Me.Caption = "Search for Loan Application"
        Case "QUOTE"
            Me.Caption = "Search for Quotations"
            ReportControlAddColumnHeader lvGrid, "Date, CustomerName,VehicleModel,ContactPerson, Telephone, Email"
            ReportControlPaintManager lvGrid
            ResizeColumnHeader lvGrid, "10, 30,20,20,10,10"
            picQuotation.Visible = True
            picQuotation.Top = 0
            picQuotation.Left = 90
        Case "CUSTOMER"
            Me.Caption = "Search for Customer"
            ReportControlAddColumnHeader lvGrid, "CustomerName, Address, Phone(s), Email "
            ReportControlPaintManager lvGrid
            ResizeColumnHeader lvGrid, "40, 20, 20,20"
            picCustomer.Visible = True
            picCustomer.Top = 0
            picCustomer.Left = 90
        Case "PO"
            Me.Caption = "Search for Purchase Order"
            ReportControlAddColumnHeader lvGrid, "DateOrdered, PONO, Description, ModelYear"
            ReportControlPaintManager lvGrid
            ResizeColumnHeader lvGrid, "20, 10, 30,10"
            picPO.Visible = True
            picPO.Top = 0
            picPO.Left = 90
        Case "MRR"
            Me.Caption = "Search for Vehicle"
            ReportControlAddColumnHeader lvGrid, "MODEL, Description, ModelYear,CSNO"
            ReportControlPaintManager lvGrid
            ResizeColumnHeader lvGrid, "20, 40, 20,20"
            picMRR.Visible = True
            picMRR.Top = 0
            picMRR.Left = 90
        Case "RELEASED"
            Me.Caption = "Search for Released Vehicles"
            ReportControlAddColumnHeader lvGrid, "CUSTOMER NAME, CS#, MODEL,VINO,PULLOUT,RELEASED"
            ReportControlPaintManager lvGrid
            ResizeColumnHeader lvGrid, "40, 20, 40,20,20,20"
            picRelease.Visible = True
            picRelease.Top = 0
            picRelease.Left = 90
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strCondition = vbNullString
    SEARCHFOR = vbNullString
End Sub

Private Sub lvGrid_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal ITEM As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If SEARCHFOR = "MRRRELEASED" Then
        If IsNull(Row.Record(2).Value) = True Then
            Metrics.BackColor = vbYellow
        End If
    End If
End Sub

Private Sub lvGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If lvGrid.SelectedRows.Count = 0 Then Exit Sub
    If KeyCode = vbKeyUp And Len(Trim(txtSearch)) <> 0 And lvGrid.SelectedRows.Row(0).Index = 0 Then
        On Error Resume Next
        txtSearch.SetFocus
    End If
    If KeyCode = 13 Then
        cmdSelect_Click
    End If
End Sub

Private Sub lvGrid_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal ITEM As XtremeReportControl.IReportRecordItem)
    cmdSelect_Click
End Sub

Private Sub optselect_Click(Index As Integer)
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub txtsearch_Change()
    Dim RS                                                            As ADODB.Recordset
    Dim SearchString                                                  As String
    Dim SQL                                                           As String
    Dim CaseStatements                                                As String
    Dim TagOption                                                     As String

    If picApplicant.Visible = True Then
        TagOption = GetControlTag(picApplicant)

        If LTrim(RTrim(ReplaceQuote(txtSearch))) = "" Then
            SearchString = ""
        Else
            SearchString = " AND " & TagOption & " like " & "'%" & ReplaceQuote(txtSearch) & "%'"
        End If

        SQL = "SELECT DateApplied, "
        SQL = SQL & "ApplicantName, "
        SQL = SQL & "UnitModel, "
        SQL = SQL & " SAE,"
        SQL = SQL & " ApplicationType, ID, ProspectId, APlCode FROM SIMS_vW_LoanApplications"
        SQL = SQL & " where LStatus='A' " & SearchString & "  ORDER BY DateApplied DESC "

    ElseIf picCustomer.Visible = True Then
        CaseStatements = " case " & vbCrLf & _
                       " when (LEN(phone) > 0  and LEN(mobile) >0  )then Phone  +  '/' + Mobile" & vbCrLf & _
                       " when (LEN(phone) > 0  and LEN(mobile) =0  ) then Phone" & vbCrLf & _
                       " when (LEN(phone) = 0  and LEN(mobile) >0  ) then Mobile" & vbCrLf & _
                       " Else ''" & vbCrLf & _
                       " end  As Contacts "

        TagOption = GetControlTag(picCustomer)

        If LTrim(RTrim(ReplaceQuote(txtSearch))) = "" Then
            SearchString = " WHERE CUSTID<>'999999' "
        Else
            SearchString = " WHERE CUSTID<>'999999'  AND " & TagOption & " like " & "'%" & ReplaceQuote(txtSearch) & "%'"
        End If

        SQL = "SELECT TOP 200 "
        SQL = SQL & " ACCTNAME, Address, " & CaseStatements & " , "
        SQL = SQL & " Email, Custype, CUSTID, CUSCDE from CRIS_Vw_allProfile "
        SQL = SQL & SearchString & " " & strCondition & " Order by " & TagOption

    ElseIf picProspect.Visible = True Then

        TagOption = GetControlTag(picProspect)

        If LTrim(RTrim(ReplaceQuote(txtSearch))) = "" Then
            SearchString = ""
        Else
            SearchString = " " & TagOption & " like " & "'%" & ReplaceQuote(txtSearch) & "%'  AND"
        End If
        SQL = "SELECT TOP 100  LogInitialInquiry, "
        SQL = SQL & " AcctName ,  Variant,SAE , TelePhone , Mobile, Address"
        SQL = SQL & " CUSCDE, ProspectType, ProspectID "
        SQL = SQL & " from CRIS_PROSPECTS WHERE "
        SQL = SQL & SearchString & "   STATUS <>'I' " & strCondition
        SQL = SQL & "ORDER BY LOGINITIALINQUIRY DESC"

    ElseIf picQuotation.Visible = True Then

        TagOption = GetControlTag(picQuotation)

        If LTrim(RTrim(ReplaceQuote(txtSearch))) = "" Then
            SearchString = ""
        Else
            SearchString = " WHERE " & TagOption & " like " & "'%" & ReplaceQuote(txtSearch) & "%'"
        End If

        SQL = " SELECT TOP 100 "
        SQL = SQL & " CQ.QuotationDate, CP.AcctName, "
        SQL = SQL & " CQ.ModelDescript, CP.ContactPerson,"
        SQL = SQL & " CP.Telephone , CP.Email ,CQ.LOGID"
        SQL = SQL & " FROM      CRIS_Quotation   "
        SQL = SQL & " CQ INNER JOIN CRIS_Prospects CP ON CQ.ProspectID = CP.ProspectID  "
        SQL = SQL & " " & SearchString & " ORDER BY CQ.QuotationDate DESC"
    ElseIf picPO.Visible = True Then
        '"DateOrdered, PONO, Model, ModelCode, Description, ModelYear,CustomerName,Source"
        TagOption = GetControlTag(picPO)

        If LTrim(RTrim(ReplaceQuote(txtSearch))) = "" Then
            SearchString = ""
        Else
            SearchString = " AND " & TagOption & " like " & "'%" & ReplaceQuote(txtSearch) & "%'"
        End If




        SQL = "SELECT SMIS_PO.DateOrdered,  "
        SQL = SQL & "SMIS_PO.PO_NO,  "
        SQL = SQL & "SMIS_PO.ModelDescript,  "
        SQL = SQL & "SMIS_PO.ModelYear, "
        SQL = SQL & "SMIS_PO.Source, "
        SQL = SQL & "SMIS_PO.ID "
        SQL = SQL & "FROM          "
        SQL = SQL & "SMIS_PO "
        SQL = SQL & "WHERE ISNULL(STATUS,'')='P' AND ISDATE(DateReceived)=0 "
        SQL = SQL & " " & SearchString & " ORDER BY " & TagOption & " DESC"

    ElseIf picMRR.Visible = True Then

        TagOption = GetControlTag(picMRR)

        If LTrim(RTrim(ReplaceQuote(txtSearch))) = "" Then
            SearchString = ""
        Else
            SearchString = " AND " & TagOption & " Like " & " '" & ReplaceQuote(txtSearch) & "%'"
        End If

        SQL = "SELECT SMIS_MRRINV.Model,  "
        SQL = SQL & "SMIS_MRRINV.Descript,  "
        SQL = SQL & "SMIS_MRRINV.Yeer,  "
        SQL = SQL & "SMIS_MRRINV.ignkey, "
        SQL = SQL & "SMIS_MRRINV.ID "
        SQL = SQL & "FROM          "
        SQL = SQL & "SMIS_MRRINV "
        SQL = SQL & "WHERE STATUS<>'P'  "
        SQL = SQL & " " & SearchString & " ORDER BY " & TagOption

    ElseIf picRelease.Visible = True Then
        TagOption = GetControlTag(picRelease)
        If LTrim(RTrim(ReplaceQuote(txtSearch))) = "" Then
            SearchString = ""
        Else
            SearchString = " AND " & TagOption & " Like " & " '%" & ReplaceQuote(txtSearch) & "%'"
        End If
        SQL = " SELECT CUSTNAME,"
        SQL = SQL & " IGNKEY_NO,"
        SQL = SQL & " UPPER(SMIS_MRRINV.MAKE + ISNULL(' ' + SMIS_MRRINV.YEER ,'')+ ISNULL(' ' + SMIS_MRRINV.MODEL ,'')),"
        SQL = SQL & " SMIS_SALESORDER.VI_NO,"
        SQL = SQL & " SMIS_MRRINV.DATERELEASED,"
        SQL = SQL & " PULLOUTDATE"
        SQL = SQL & " FROM SMIS_SALESORDER LEFT OUTER JOIN SMIS_MRRINV ON IGNKEY_NO=IGNKEY "
        SQL = SQL & " WHERE SMIS_SALESORDER.SOSTATUS<>'C' AND SMIS_SALESORDER.STATUS='P' AND SMIS_SALESORDER.VI_NO NOT IN (SELECT VI_NO FROM SMIS_COMMISSION) "
        SQL = SQL & " " & SearchString & " ORDER BY " & TagOption

    End If
    Set RS = gconDMIS.Execute(SQL)
    
    flex_FillReportView RS, lvGrid, False
    If lvGrid.Records.Count = 0 Then
        cmdSelect.Enabled = False
    Else
        cmdSelect.Enabled = True
    End If

End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)

    If Trim(txtSearch) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If

    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then

        If lvGrid.Enabled = True Then: lvGrid.SetFocus
    End If

    If KeyCode = vbKeyEscape Then Unload Me

End Sub


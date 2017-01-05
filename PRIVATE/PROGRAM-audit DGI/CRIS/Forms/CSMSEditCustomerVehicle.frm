VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Object = "{A9046457-E246-455F-A58F-D670C44E8BEA}#2.0#0"; "wizFlexCracker.ocx"
Begin VB.Form frmCSMSEditCustomerVehicle 
   BackColor       =   &H8000000C&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Customer Vehicle Maintenance"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13110
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   13110
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8400
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox picSearch 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   -30
      ScaleHeight     =   885
      ScaleWidth      =   8115
      TabIndex        =   8
      Top             =   7590
      Width           =   8115
      Begin VB.OptionButton optEndUser 
         BackColor       =   &H8000000C&
         Caption         =   "End User"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   5820
         TabIndex        =   17
         Top             =   30
         Width           =   1425
      End
      Begin VB.OptionButton optSellingDealer 
         BackColor       =   &H8000000C&
         Caption         =   "Selling Dealer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   4350
         TabIndex        =   16
         Top             =   30
         Width           =   1485
      End
      Begin VB.OptionButton optCSNo 
         BackColor       =   &H8000000C&
         Caption         =   "CS#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   2730
         TabIndex        =   15
         Top             =   30
         Width           =   735
      End
      Begin VB.OptionButton optModel 
         BackColor       =   &H8000000C&
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   3450
         TabIndex        =   14
         Top             =   30
         Width           =   1095
      End
      Begin VB.OptionButton optCustomername 
         BackColor       =   &H8000000C&
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   990
         TabIndex        =   13
         Top             =   30
         Value           =   -1  'True
         Width           =   1665
      End
      Begin VB.OptionButton optPlate 
         BackColor       =   &H8000000C&
         Caption         =   "Plate#"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   210
         Left            =   60
         TabIndex        =   12
         Top             =   30
         Width           =   825
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   405
         Left            =   4590
         Picture         =   "CSMSEditCustomerVehicle.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   525
      End
      Begin VB.TextBox txtSearch 
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
         Left            =   1650
         TabIndex        =   9
         Top             =   360
         Width           =   2925
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         Caption         =   "Search Key Word"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   10
         Top             =   420
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   12300
      MouseIcon       =   "CSMSEditCustomerVehicle.frx":047F
      MousePointer    =   99  'Custom
      Picture         =   "CSMSEditCustomerVehicle.frx":05D1
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Cancel"
      Top             =   7650
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   795
      Left            =   11580
      MouseIcon       =   "CSMSEditCustomerVehicle.frx":090F
      MousePointer    =   99  'Custom
      Picture         =   "CSMSEditCustomerVehicle.frx":0A61
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Save Customer Vehicle Information"
      Top             =   7650
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   795
      Left            =   10860
      MouseIcon       =   "CSMSEditCustomerVehicle.frx":0DB1
      MousePointer    =   99  'Custom
      Picture         =   "CSMSEditCustomerVehicle.frx":0F03
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Print this Record"
      Top             =   7650
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
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
      Height          =   1065
      Left            =   3983
      ScaleHeight     =   1035
      ScaleWidth      =   5115
      TabIndex        =   2
      Top             =   3698
      Visible         =   0   'False
      Width           =   5145
      Begin MSComctlLib.ProgressBar prgCusVeh 
         Height          =   405
         Left            =   60
         TabIndex        =   4
         Top             =   420
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label labCusVeh 
         Caption         =   "Updating Records..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   3
         Top             =   90
         Width           =   4725
      End
   End
   Begin FlexCell.Grid grdCusVeh 
      Height          =   7575
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   13361
      BackColor2      =   16701142
      BackColorBkg    =   -2147483645
      Cols            =   20
      DefaultFontSize =   8.25
      DisplayRowIndex =   -1  'True
      Rows            =   2
   End
   Begin wizFlexCracker.wizFlexCrack wizFlexCrack1 
      Height          =   3765
      Left            =   1500
      TabIndex        =   1
      Top             =   2070
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   6641
   End
End
Attribute VB_Name = "frmCSMSEditCustomerVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
If Function_Access(LOGID, "Acess_PRINT", "CUSTOMER VEHICLE") = False Then Exit Sub
CrystalReport1.Formulas(0) = "companyname='" & COMPANY_NAME & "'"
CrystalReport1.Formulas(1) = "COMPANYADDRESS='" & COMPANY_ADDRESS & "'"
    PrintSQLReport CrystalReport1, CSMS_REPORT_PATH & "cusvehreports.rpt", "", DMIS_REPORT_Connection, 1
End Sub

Private Sub cmdSave_Click()
    If Function_Access(LOGID, "Acess_ADD", "CUSTOMER VEHICLE") = False Then Exit Sub
    
    Dim Harry                           As Integer
    Screen.MousePointer = 11: Picture1.Visible = True: prgCusVeh.Value = 0
    grdCusVeh.AutoRedraw = False
    With grdCusVeh
        For Harry = 1 To .Rows - 1
            gconDMIS.Execute ("Update CSMS_CusVeh Set " & _
                            " CUSCDE = " & N2Str2Null(SetCustomerCode(.Cell(Harry, 1).Text)) & "," & _
                            " Plate_No = " & N2Str2Null(.Cell(Harry, 2).Text) & "," & _
                            " VCond_No = " & N2Str2Null(.Cell(Harry, 3).Text) & "," & _
                            " ClrCde = " & N2Str2Null(SetColorCode(.Cell(Harry, 4).Text)) & "," & _
                            " Yer = " & N2Str2Null(.Cell(Harry, 5).Text) & "," & _
                            " Make = " & N2Str2Null(.Cell(Harry, 6).Text) & "," & _
                            " Model = " & N2Str2Null(.Cell(Harry, 7).Text) & "," & _
                            " Description = " & N2Str2Null(.Cell(Harry, 8).Text) & "," & _
                            " Engine = " & N2Str2Null(.Cell(Harry, 9).Text) & "," & _
                            " KMReading = " & N2Str2Null(.Cell(Harry, 10).Text) & "," & _
                            " ProdNo = " & N2Str2Null(.Cell(Harry, 11).Text) & "," & _
                            " TIN_Number = " & N2Str2Null(.Cell(Harry, 12).Text) & "," & _
                            " D_Sold = " & N2Str2Null(.Cell(Harry, 13).Text) & "," & _
                            " InvoiceNo = " & N2Str2Null(.Cell(Harry, 14).Text) & "," & _
                            " War_Cert = " & N2Str2Null(.Cell(Harry, 15).Text) & "," & _
                            " Del_Date = " & N2Str2Null(.Cell(Harry, 16).Text) & "," & _
                            " Selling_Dealer = " & N2Str2Null(SetSellingDealerCode(.Cell(Harry, 17).Text)) & "," & _
                            " EndUser = " & N2Str2Null(SetCustomerCode(.Cell(Harry, 18).Text)) & _
                            " Where ID = " & .Cell(Harry, 19).Text)
            DoEvents
            prgCusVeh.Value = (Harry / (.Rows - 1)) * 100
            labCusVeh.Caption = "Updating Records... " & Int(prgCusVeh.Value) & "% Completed"
        Next
    End With
    grdCusVeh.AutoRedraw = True
    grdCusVeh.Refresh
    Screen.MousePointer = 0: Picture1.Visible = False
    MsgBox "Records Successfully Updated.", vbInformation, "Updated"
End Sub

Private Sub cmdSearch_Click()
txtSearch_KeyPress 13
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    InitGrid
    Screen.MousePointer = 11
    Me.Caption = "Displaying Vehicle Info"
    StoreGridDetails
    Me.Caption = "Edit Vehicle Info"
    Screen.MousePointer = 0
End Sub

Sub StoreGridDetails()
    Dim kim                             As Integer
    Dim rsCusVeh                        As ADODB.Recordset
    Set rsCusVeh = New ADODB.Recordset
    Set rsCusVeh = gconDMIS.Execute("Select * from CSMS_CusVeh Order by Plate_No ASC")
    If Not rsCusVeh.EOF And Not rsCusVeh.BOF Then
        rsCusVeh.MoveFirst: kim = 0
        grdCusVeh.AutoRedraw = False
        Do While Not rsCusVeh.EOF
            kim = kim + 1
            grdCusVeh.AddItem SetCustomerName(Null2String(rsCusVeh!CUSCDE)) & Chr(9) & _
                              Null2String(rsCusVeh!PLATE_NO) & Chr(9) & _
                              Null2String(rsCusVeh!VCOND_NO) & Chr(9) & _
                              SetColorDesc(Null2String(rsCusVeh!ClrCde)) & Chr(9) & _
                              Null2String(rsCusVeh!Yer) & Chr(9) & _
                              Null2String(rsCusVeh!Make) & Chr(9) & _
                              Null2String(rsCusVeh!MODEL) & Chr(9) & _
                              Null2String(rsCusVeh!Description) & Chr(9) & _
                              Null2String(rsCusVeh!Engine) & Chr(9) & _
                              Null2String(rsCusVeh!KMReading) & Chr(9) & _
                              Null2String(rsCusVeh!ProdNo) & Chr(9) & _
                              Null2String(rsCusVeh!TIN_Number) & Chr(9) & _
                              Null2String(rsCusVeh!D_Sold) & Chr(9) & _
                              Null2String(rsCusVeh!InvoiceNo) & Chr(9) & _
                              Null2String(rsCusVeh!War_Cert) & Chr(9) & _
                              Null2String(rsCusVeh!Del_Date) & Chr(9) & _
                              SetSellingDealerName(Null2String(rsCusVeh!Selling_Dealer)) & Chr(9) & _
                              SetCustomerName(Null2String(rsCusVeh!EndUser)) & Chr(9) & _
                              rsCusVeh!ID
            DoEvents
            If kim = 1 Then grdCusVeh.RemoveItem 1
            rsCusVeh.MoveNext
        Loop
        grdCusVeh.AutoRedraw = True
        grdCusVeh.Refresh
    End If
End Sub
Sub SearchGrid()
  Dim Harry                           As Integer
    Screen.MousePointer = 11: Picture1.Visible = True: prgCusVeh.Value = 0
    Dim kim                             As Integer
    Dim rsCusVeh                        As ADODB.Recordset
    Dim maxval As Integer
    Set rsCusVeh = New ADODB.Recordset

    If optCSNo.Value = True Then
        Call rsCusVeh.Open("Select * from CSMS_CusVeh where vcond_no like '" & Replace(txtSearch.Text, "'", "") & "%' Order by vcond_no ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly)
    ElseIf optCustomername.Value = True Then
        Call rsCusVeh.Open("Select * from CSMS_CusVeh where NIYM like '" & Replace(txtSearch.Text, "'", "") & "%' Order by NIYM ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly)
    ElseIf optModel.Value = True Then
        Call rsCusVeh.Open("Select * from CSMS_CusVeh where model like '" & Replace(txtSearch.Text, "'", "") & "%' Order by model ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly)
    ElseIf optPlate.Value = True Then
        Call rsCusVeh.Open("Select * from CSMS_CusVeh where PLATE_NO like '" & Replace(txtSearch.Text, "'", "") & "%' Order by PLATE_NO ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly)
    ElseIf optSellingDealer.Value = True Then
        Call rsCusVeh.Open("Select * from CSMS_CusVeh where SELLING_DEALER like '" & Replace(txtSearch.Text, "'", "") & "%' Order by SELLING_DEALER ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly)
    Else
        Call rsCusVeh.Open("Select CSMS_CusVeh.* from CSMS_CusVeh INNER JOIN ALL_CUSTOMER_TABLE ON ALL_CUSTOMER_TABLE.CUSCDE = CSMS_CusVeh.ENDUSER where ALL_CUSTOMER_TABLE ACCTNAME LIKE '" & Replace(txtSearch.Text, "'", "") & "%' Order by SELLING_DEALER ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly)
    End If
        grdCusVeh.Rows = 1
        maxval = rsCusVeh.RecordCount
        
    If Not rsCusVeh.EOF And Not rsCusVeh.BOF Then
        rsCusVeh.MoveFirst
        kim = 0: prgCusVeh.Max = maxval
        grdCusVeh.AutoRedraw = False
        Do While Not rsCusVeh.EOF
            kim = kim + 1
            grdCusVeh.AddItem SetCustomerName(Null2String(rsCusVeh!CUSCDE)) & Chr(9) & _
                              Null2String(rsCusVeh!PLATE_NO) & Chr(9) & _
                              Null2String(rsCusVeh!VCOND_NO) & Chr(9) & _
                              SetColorDesc(Null2String(rsCusVeh!ClrCde)) & Chr(9) & _
                              Null2String(rsCusVeh!Yer) & Chr(9) & _
                              Null2String(rsCusVeh!Make) & Chr(9) & _
                              Null2String(rsCusVeh!MODEL) & Chr(9) & _
                              Null2String(rsCusVeh!Description) & Chr(9) & _
                              Null2String(rsCusVeh!Engine) & Chr(9) & _
                              Null2String(rsCusVeh!KMReading) & Chr(9) & _
                              Null2String(rsCusVeh!ProdNo) & Chr(9) & _
                              Null2String(rsCusVeh!TIN_Number) & Chr(9) & _
                              Null2String(rsCusVeh!D_Sold) & Chr(9) & _
                              Null2String(rsCusVeh!InvoiceNo) & Chr(9) & _
                              Null2String(rsCusVeh!War_Cert) & Chr(9) & _
                              Null2String(rsCusVeh!Del_Date) & Chr(9) & _
                              SetSellingDealerName(Null2String(rsCusVeh!Selling_Dealer)) & Chr(9) & _
                              SetCustomerName(Null2String(rsCusVeh!EndUser)) & Chr(9) & _
                              rsCusVeh!ID, False
            
            prgCusVeh.Value = kim
            labCusVeh.Caption = "Storing Records... " & Round((prgCusVeh.Value / maxval) * 100, 2) & " % Completed"
            DoEvents
            rsCusVeh.MoveNext
        Loop
        grdCusVeh.AutoRedraw = True: grdCusVeh.Refresh
    End If
    If grdCusVeh.Rows = 1 Then: cmdSave.Enabled = False: Else: cmdSave.Enabled = True
Screen.MousePointer = 0: Picture1.Visible = False: prgCusVeh.Value = 0
End Sub
Sub InitGrid()
    With grdCusVeh
    
        .Cell(0, 1).Text = "Customer Name"
        .Cell(0, 2).Text = "Plate No."
        .Cell(0, 3).Text = "Cond. No."
        .Cell(0, 4).Text = "Color"
        .Cell(0, 5).Text = "Year"
        .Cell(0, 6).Text = "Make"
        .Cell(0, 7).Text = "Model"
        .Cell(0, 8).Text = "Description"
        .Cell(0, 9).Text = "Engine"
        .Cell(0, 10).Text = "KM Rdg"
        .Cell(0, 11).Text = "Prod. #"
        .Cell(0, 12).Text = "TIN No."
        .Cell(0, 13).Text = "Date Sold"
        .Cell(0, 14).Text = "Invoice#"
        .Cell(0, 15).Text = "Warranty Cert."
        .Cell(0, 16).Text = "Date Delvrd"
        .Cell(0, 17).Text = "Selling Dealer"
        .Cell(0, 18).Text = "End User"
        .Cell(0, 19).Text = "ID"

        .Column(0).Width = 0
        .Column(1).Width = 270
        .Column(2).Width = 55
        .Column(2).MaxLength = 6
        .Column(3).Width = 60
        .Column(4).Width = 120
        .Column(5).Width = 50
        .Column(6).Width = 80
        .Column(7).Width = 100
        .Column(8).Width = 140
        .Column(9).Width = 80
        .Column(10).Width = 50
        .Column(11).Width = 70
        .Column(12).Width = 70
        .Column(13).Width = 70
        .Column(14).Width = 60
        .Column(15).Width = 70
        .Column(16).Width = 70
        .Column(17).Width = 200
        .Column(18).Width = 200
        .Column(19).Width = 0

        .Column(1).CellType = cellComboBox
        .ComboBox(1).Locked = True
        .ComboBox(1).Font = "TAHOMA"
        .Column(4).CellType = cellComboBox
        .ComboBox(4).Locked = True
        .ComboBox(4).Font = "TAHOMA"
        .Column(6).CellType = cellComboBox
        .ComboBox(6).Locked = True
        .ComboBox(6).Font = "TAHOMA"
        .Column(7).CellType = cellComboBox
        .ComboBox(7).Locked = True
        .ComboBox(7).Font = "TAHOMA"
        .Column(13).CellType = cellCalendar
        .Column(16).CellType = cellCalendar
        .Column(17).CellType = cellComboBox
        .ComboBox(17).Locked = True
        .ComboBox(17).Font = "TAHOMA"
        .Column(18).CellType = cellComboBox
        .ComboBox(18).Locked = True
        .ComboBox(18).Font = "TAHOMA"

        grdCusVeh.DefaultFont = "tahoma"

        With grdCusVeh.ComboBox(1)
            Dim rsCustomer              As ADODB.Recordset
            Set rsCustomer = New ADODB.Recordset
            Set rsCustomer = gconDMIS.Execute("Select * from ALL_Customer Order by AcctName asc")
            If Not rsCustomer.EOF And Not rsCustomer.BOF Then
                rsCustomer.MoveFirst
                Do While Not rsCustomer.EOF
                    .AddItem Null2String(rsCustomer!AcctName)
                    grdCusVeh.ComboBox(18).AddItem Null2String(rsCustomer!AcctName)
                    rsCustomer.MoveNext
                Loop
            End If
            Set rsCustomer = Nothing
        End With

        With grdCusVeh.ComboBox(4)
            Dim rsColor                 As ADODB.Recordset
            Set rsColor = New ADODB.Recordset
            Set rsColor = gconDMIS.Execute("Select distinct Color_Desc  from ALL_Color Order by Color_Desc asc")
            If Not rsColor.EOF And Not rsColor.BOF Then
                rsColor.MoveFirst
                Do While Not rsColor.EOF
                    .AddItem Null2String(rsColor!Color_Desc)
                    rsColor.MoveNext
                Loop
            End If
            Set rsColor = Nothing
        End With

        With grdCusVeh.ComboBox(6)
            Dim rsMake                  As ADODB.Recordset
            Set rsMake = New ADODB.Recordset
            Set rsMake = gconDMIS.Execute("Select distinct make from ALL_Make Order by Make asc")
            If Not rsMake.EOF And Not rsMake.BOF Then
                rsMake.MoveFirst
                Do While Not rsMake.EOF
                    .AddItem UCase(Null2String(rsMake!Make))
                    rsMake.MoveNext
                Loop
            End If
            Set rsColor = Nothing
        End With

        With grdCusVeh.ComboBox(7)
            Dim rsModel                 As ADODB.Recordset
            Set rsModel = New ADODB.Recordset
            Set rsModel = gconDMIS.Execute("Select distinct [MODEL] from CSMS_MODELS Order BY 1")
            If Not rsModel.EOF And Not rsModel.BOF Then
                rsModel.MoveFirst
                Do While Not rsModel.EOF
                    .AddItem UCase(Null2String(rsModel!MODEL))
                    rsModel.MoveNext
                Loop
            End If
            Set rsModel = Nothing
        End With

        With grdCusVeh.ComboBox(17)
            Dim rsSellingDealer         As ADODB.Recordset
            Set rsSellingDealer = New ADODB.Recordset
            Set rsSellingDealer = gconDMIS.Execute("Select * from CSMS_SellingDealer order by DealerName asc")
            If Not rsSellingDealer.EOF And Not rsSellingDealer.BOF Then
                rsSellingDealer.MoveFirst
                Do While Not rsSellingDealer.EOF
                    .AddItem rsSellingDealer!DealerName
                    rsSellingDealer.MoveNext
                Loop
            End If
            Set rsSellingDealer = Nothing
        End With
    End With
End Sub

Function SetCustomerName(XXX As String) As String
    Dim rsCustomer                      As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from ALL_Customer Where CUSCDE = '" & XXX & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerName = Null2String(rsCustomer!AcctName)
    End If
End Function

Function SetCustomerCode(XXX As String) As String
    Dim rsCustomer                      As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from ALL_Customer Where AcctName = " & N2Str2Null(XXX))
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerCode = Null2String(rsCustomer!CUSCDE)
    End If
End Function

Function SetColorDesc(XXX As String) As String
    Dim rsCustomer                      As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from ALL_Color Where Color_Code = " & N2Str2Null(XXX))
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetColorDesc = Null2String(rsCustomer!Color_Desc)
    End If
End Function

Function SetColorCode(XXX As String) As String
    Dim rsCustomer                      As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from ALL_Color Where Color_Desc = '" & XXX & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetColorCode = Null2String(rsCustomer!Color_Code)
    End If
End Function

Function SetSellingDealerName(XXX As String) As String
    Dim rsCustomer                      As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from CSMS_SellingDealer Where DealerCode = '" & XXX & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetSellingDealerName = Null2String(rsCustomer!DealerName)
    End If
End Function

Function SetSellingDealerCode(XXX As String) As String
    Dim rsCustomer                      As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from CSMS_SellingDealer Where DealerName = '" & XXX & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetSellingDealerCode = Null2String(rsCustomer!DealerCode)
    End If
End Function



Private Sub optCSNo_Click()
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub optCustomername_Click()
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub optEndUser_Click()
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub optModel_Click()
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub optPlate_Click()
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub optSellingDealer_Click()
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(LTrim(RTrim(txtSearch))) > 0 Then
            SearchGrid
        End If
    End If
End Sub

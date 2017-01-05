VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmSMIS_Inquiry_CustomerSalesHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Customer and Vehicles Sales History inquiry"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12270
   Icon            =   "CustomerSalesHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   12270
   Begin MSComctlLib.ListView lstAllCustomer 
      Height          =   2475
      Left            =   30
      TabIndex        =   0
      Top             =   810
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   4366
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
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
      MouseIcon       =   "CustomerSalesHistory.frx":030A
      NumItems        =   0
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "All Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10770
      TabIndex        =   5
      Top             =   210
      Width           =   1425
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "W/Transation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   3
      Top             =   210
      Width           =   1365
   End
   Begin XtremeSuiteControls.TabControl MyTab 
      Height          =   4095
      Left            =   30
      TabIndex        =   1
      Top             =   3330
      Width           =   12135
      _Version        =   655364
      _ExtentX        =   21405
      _ExtentY        =   7223
      _StockProps     =   64
      Appearance      =   9
      Color           =   4
      PaintManager.Layout=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.FixedTabWidth=   150
      PaintManager.MinTabWidth=   100
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Vehicle Invoice"
      Item(0).Tooltip =   "Vehicle Invoice"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lvVehicleSalesInvoice"
      Item(1).Caption =   "Customer Vehicle Info"
      Item(1).Tooltip =   "Customer Vehicle Information"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lvCustomerVehicles"
      Begin MSComctlLib.ListView lvVehicleSalesInvoice 
         Height          =   3645
         Left            =   -69940
         TabIndex        =   2
         Top             =   390
         Visible         =   0   'False
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   6429
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
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
         MouseIcon       =   "CustomerSalesHistory.frx":046C
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvCustomerVehicles 
         Height          =   3645
         Left            =   60
         TabIndex        =   14
         Top             =   360
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   6429
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
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
         MouseIcon       =   "CustomerSalesHistory.frx":05CE
         NumItems        =   0
      End
   End
   Begin VB.Frame fraAllCustomer 
      Caption         =   "All Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   9195
      Begin VB.OptionButton Otp 
         Caption         =   "By Customer Name"
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
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Otp 
         Caption         =   "By Contact Person"
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
         Index           =   1
         Left            =   2130
         TabIndex        =   9
         Top             =   240
         Width           =   1845
      End
      Begin VB.OptionButton Otp 
         Caption         =   "By Cuscde"
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
         Index           =   2
         Left            =   4110
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.TextBox txtSearchKey_All 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5370
         TabIndex        =   7
         Top             =   210
         Width           =   3705
      End
   End
   Begin VB.Frame fraActiveCustomer 
      Caption         =   "W/Transaction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   30
      TabIndex        =   4
      Top             =   30
      Width           =   9195
      Begin VB.TextBox txtSearchKey_Active 
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
         Left            =   4410
         TabIndex        =   13
         Top             =   270
         Width           =   4665
      End
      Begin VB.OptionButton optSearchActiveCode 
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2370
         TabIndex        =   12
         Top             =   300
         Width           =   1215
      End
      Begin VB.OptionButton optSearchActiveCustomerName 
         Caption         =   "By Customer Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   11
         Top             =   300
         Value           =   -1  'True
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmSMIS_Inquiry_CustomerSalesHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TheCuscde                                                         As String

Sub Fill_ActiveCustomerSearch()
    Dim SQL                                                           As String
    Dim RS                                                            As New ADODB.Recordset
    Dim Keyword                                                       As String
    On Error GoTo ErrorCode:
    lstAllCustomer.Enabled = False



    SQL = "SELECT CUSCDE, CUSTOMERNAME, ADDRESS, CONTACTPERSON, EMAIL, PHONE, MOBILE FROM CRIS_vw_AllProfile  where CUSCDE IN (SELECT CODE FROM SMIS_SALESORDER)"
    Keyword = RTrim(LTrim(txtSearchKey_Active.Text))


    If optSearchActiveCustomerName.Value = True Then
        SQL = SQL & "  AND  CUSTOMERNAME LIKE '" & ReplaceQuote(Keyword) & "%'"
    End If
    If optSearchActiveCode.Value = True Then
        SQL = SQL & "  AND  CUSCDE LIKE '" & ReplaceQuote(Keyword) & "%'"
    End If
    Set RS = gconDMIS.Execute(SQL)


    If Not RS.EOF And Not RS.BOF Then
        lstAllCustomer.Enabled = True
    End If


    flex_FillListView RS, lstAllCustomer, True, True

    LV_AutoSizeColumn lstAllCustomer



    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Sub Fill_AllCustomerSearch()
    Dim SQL                                                           As String
    Dim RS                                                            As New ADODB.Recordset
    Dim Keyword                                                       As String



    lstAllCustomer.Enabled = False

    On Error GoTo ErrorCode:



    SQL = "SELECT CUSCDE, CUSTOMERNAME, ADDRESS, CONTACTPERSON, EMAIL, PHONE, MOBILE FROM CRIS_vw_AllProfile"

    Keyword = RTrim(LTrim(txtSearchKey_All.Text))




    If Otp(0).Value = True Then
        SQL = SQL & " WHERE CUSTOMERNAME LIKE '" & ReplaceQuote(Keyword) & "%'"
    End If

    If Otp(1).Value = True Then
        SQL = SQL & " WHERE CONTACTPERSON LIKE'" & ReplaceQuote(Keyword) & "%'"
    End If

    If Otp(2).Value = True Then
        SQL = SQL & " WHERE cuscde like '" & ReplaceQuote(Keyword) & "%'"
    End If



    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        lstAllCustomer.Enabled = True
    End If


    flex_FillListView RS, lstAllCustomer, True, True





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Sub FillVehicle()
    Dim RS                                                            As New ADODB.Recordset
    Dim SQL                                                           As String

    SQL = " SELECT"
    SQL = SQL & " PLATE_NO AS [PLATE NUMBER],"
    SQL = SQL & " (SELECT TOP 1 COLOR_DESC FROM ALL_COLOR WHERE COLOR_CODE=CLRCDE)AS COLOR ,"
    SQL = SQL & " ISNULL(MAKE + ' ' ,'') + ISNULL(YER + ' ','') + ISNULL((SELECT TOP 1  [DESCRIPTION] FROM ALL_MODELCODE WHERE CODE=MODEL ),'') ,"
    SQL = SQL & " VIN VINNO, VCOND_NO CSNO,"
    SQL = SQL & " ENGINE, PRODNO,KMREADING,"
    SQL = SQL & " INVOICENO , SELLING_DEALER AS [SELLING DEALER],"
    SQL = SQL & " WAR_CERT FROM CSMS_CUSVEH WHERE cuscde='" & TheCuscde & "'"

    Set RS = gconDMIS.Execute(SQL)
    flex_FillListView RS, lvCustomerVehicles
    Set RS = Nothing
End Sub

Sub FillVehicleInvoice()
    Dim RS                                                            As ADODB.Recordset
    Dim SQL                                                           As String

    lvVehicleSalesInvoice.Enabled = False

    SQL = "SELECT  DateReleased , InvoicedDate, VI_NO [VI No],"
    SQL = SQL & " ModelDescription [Description], ProdNo [P#], EngineNo [E#], FrameNo [F#], Vino [VIN#], Plate_No, IGNKEY_NO [CS#], Color, Type,"
    SQL = SQL & " Term , FinancingCo, SalesAE, SalesApproved, Insured, Certific8, Status "
    SQL = SQL & " FROM SMIS_SALESORDER where VI_NO is not Null and STATUS<>'C' AND CODE=" & N2Str2Null(TheCuscde)

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        lvVehicleSalesInvoice.Enabled = True
    End If

    flex_FillListView RS, lvVehicleSalesInvoice, True, True

    Set RS = Nothing


End Sub

Sub ConfigGrids()
    AddColumnHeader "Plate, Color, Model, V#,CS#, E#, P#, KM, INV#, Selling Dealer, Warranty", lvCustomerVehicles
    ResizeColumnHeader lvCustomerVehicles, "8,10,20,10,10,10,10,6,8"


End Sub

Private Sub CmdAll_Click()
    fraAllCustomer.Visible = True
    fraActiveCustomer.Visible = False
    Fill_AllCustomerSearch
    If lstAllCustomer.ListItems.Count > 0 Then
        lstAllCustomer.ListItems(1).Selected = True
        lstAllCustomer.ListItems(1).EnsureVisible
        lstAllCustomer_ItemClick lstAllCustomer.SelectedItem
    Else
        lvVehicleSalesInvoice.ListItems.Clear

        lvCustomerVehicles.ListItems.Clear
    End If
    On Error Resume Next
    txtSearchKey_All.SetFocus
End Sub

Private Sub cmdTrans_Click()
    Fill_ActiveCustomerSearch
    fraAllCustomer.Visible = False
    fraActiveCustomer.Visible = True
    If lstAllCustomer.ListItems.Count > 0 Then
        lstAllCustomer.ListItems(1).Selected = True
        lstAllCustomer.ListItems(1).EnsureVisible
        lstAllCustomer_ItemClick lstAllCustomer.SelectedItem
    Else
        lvVehicleSalesInvoice.ListItems.Clear
        lvCustomerVehicles.ListItems.Clear
    End If
    On Error Resume Next
    txtSearchKey_Active.SetFocus

End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    ConfigGrids
    cmdTrans_Click
End Sub

Private Sub lstAllCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    TheCuscde = lstAllCustomer.SelectedItem.SubItems(1)
    Call FillVehicleInvoice
    Call FillVehicle

End Sub

Private Sub Otp_Click(Index As Integer)
    On Error Resume Next
    txtSearchKey_All.SetFocus
End Sub

Private Sub txtSearchKey_Active_Change()
    Call Fill_ActiveCustomerSearch
    If lstAllCustomer.ListItems.Count > 0 Then
        lstAllCustomer.ListItems(1).Selected = True
        lstAllCustomer.ListItems(1).EnsureVisible
        lstAllCustomer_ItemClick lstAllCustomer.SelectedItem
    Else
        lvVehicleSalesInvoice.ListItems.Clear
        lvCustomerVehicles.ListItems.Clear
    End If

End Sub

Private Sub txtSearchKey_All_Change()
    Fill_AllCustomerSearch
    If lstAllCustomer.ListItems.Count > 0 Then
        lstAllCustomer.ListItems(1).Selected = True
        lstAllCustomer.ListItems(1).EnsureVisible
        lstAllCustomer_ItemClick lstAllCustomer.SelectedItem
    Else
        lvVehicleSalesInvoice.ListItems.Clear
        lvCustomerVehicles.ListItems.Clear
    End If
End Sub


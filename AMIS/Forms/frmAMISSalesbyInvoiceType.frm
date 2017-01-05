VERSION 5.00
Begin VB.Form frmAMIS_SalesbyInvoiceType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales By Invoice Type"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   Icon            =   "frmAMISSalesbyInvoiceType.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1320
      MouseIcon       =   "frmAMISSalesbyInvoiceType.frx":27A2
      MousePointer    =   99  'Custom
      Picture         =   "frmAMISSalesbyInvoiceType.frx":28F4
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Process"
      Top             =   1230
      Width           =   780
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   2085
      MouseIcon       =   "frmAMISSalesbyInvoiceType.frx":2B8F
      MousePointer    =   99  'Custom
      Picture         =   "frmAMISSalesbyInvoiceType.frx":2CE1
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Cancel"
      Top             =   1230
      Width           =   780
   End
   Begin VB.PictureBox pixPartsInvoice 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   360
      ScaleHeight     =   600
      ScaleWidth      =   4005
      TabIndex        =   2
      Top             =   570
      Width           =   4005
      Begin VB.OptionButton optPartsChargeInvoices 
         Caption         =   "Parts Charge Invoices"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   75
         TabIndex        =   6
         Top             =   330
         Width           =   3405
      End
      Begin VB.OptionButton optPartsCashInvoices 
         Caption         =   "Parts Cash Invoices"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   90
         TabIndex        =   5
         Top             =   60
         Value           =   -1  'True
         Width           =   3870
      End
      Begin VB.OptionButton optHyundaiPartsChargeInvoices 
         Caption         =   "Hyundai Parts Charge Invoices"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   510
         Left            =   75
         TabIndex        =   4
         Top             =   840
         Width           =   3855
      End
      Begin VB.OptionButton optHyundaiPartsCashInvoices 
         Caption         =   "Hyundai Parts Cash Invoices"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   75
         TabIndex        =   3
         Top             =   630
         Width           =   3405
      End
   End
   Begin VB.ComboBox cboSalesInvoiceType 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   330
      Left            =   2085
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   2175
   End
   Begin VB.PictureBox pixSalesInvoices 
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   330
      ScaleHeight     =   1305
      ScaleWidth      =   4005
      TabIndex        =   12
      Top             =   570
      Width           =   4005
      Begin VB.OptionButton optHyundaiVehicleSalesInvoices 
         Caption         =   "Hyundai Vehicle Sales Invoices"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   300
         Visible         =   0   'False
         Width           =   3405
      End
      Begin VB.OptionButton optVehicleSalesInvoices 
         Caption         =   "Vehicle Sales Invoices"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   75
         TabIndex        =   13
         Top             =   30
         Value           =   -1  'True
         Width           =   3870
      End
   End
   Begin VB.PictureBox pixServiceInvoices 
      BorderStyle     =   0  'None
      Height          =   1410
      Left            =   285
      ScaleHeight     =   1410
      ScaleWidth      =   4005
      TabIndex        =   7
      Top             =   540
      Width           =   4005
      Begin VB.OptionButton optHyundaiServiceCashInvoices 
         Caption         =   "Hyundai Service Cash Invoices"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   75
         TabIndex        =   11
         Top             =   630
         Visible         =   0   'False
         Width           =   3405
      End
      Begin VB.OptionButton optHyundaiServiceChargeInvoices 
         Caption         =   "Hyundai Service Charge Invoices"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   510
         Left            =   75
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.OptionButton optServiceCashInvoices 
         Caption         =   "Service Cash Invoices"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   75
         TabIndex        =   9
         Top             =   30
         Value           =   -1  'True
         Width           =   3870
      End
      Begin VB.OptionButton optServiceChargeInvoices 
         Caption         =   "Service Charge Invoices"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   75
         TabIndex        =   8
         Top             =   330
         Width           =   3405
      End
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Invoice Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   15
      TabIndex        =   1
      Top             =   210
      Width           =   1995
   End
End
Attribute VB_Name = "frmAMIS_SalesbyInvoiceType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function FillCombo()
    cboSalesInvoiceType.AddItem "Sales Invoice"
    cboSalesInvoiceType.AddItem "Parts Invoice"
    If COMPANY_CODE = "CMC" Then
    cboSalesInvoiceType.AddItem "Materials Invoice"
    cboSalesInvoiceType.AddItem "Accessories Invoice"
    End If
    cboSalesInvoiceType.AddItem "Service Invoice"
End Function

Private Sub cboSalesInvoiceType_Click()
    If cboSalesInvoiceType.Text = "Sales Invoice" Then
        pixSalesInvoices.Visible = True
        pixPartsInvoice.Visible = False
        pixServiceInvoices.Visible = False
    ElseIf cboSalesInvoiceType.Text = "Parts Invoice" Then
        If COMPANY_CODE = "DJM" Or COMPANY_CODE = "HCA" Then
            pixPartsInvoice.Visible = False
            pixSalesInvoices.Visible = False
            pixServiceInvoices.Visible = False
            cmdOK.Value = True
        Else
            pixPartsInvoice.Visible = True
            pixSalesInvoices.Visible = False
            pixServiceInvoices.Visible = False
            
            optPartsCashInvoices.Caption = "Parts Cash Invoices"
            optPartsChargeInvoices.Caption = "Parts Charge Invoices"
        End If
    ElseIf cboSalesInvoiceType.Text = "Materials Invoice" Then
            pixPartsInvoice.Visible = True
            pixSalesInvoices.Visible = False
            pixServiceInvoices.Visible = False
            
            optPartsCashInvoices.Caption = "Materials Cash Invoices"
            optPartsChargeInvoices.Caption = "Materials Charge Invoices"
    ElseIf cboSalesInvoiceType.Text = "Accessories Invoice" Then
            pixPartsInvoice.Visible = True
            pixSalesInvoices.Visible = False
            pixServiceInvoices.Visible = False
            
            optPartsCashInvoices.Caption = "Accessories Cash Invoices"
            optPartsChargeInvoices.Caption = "Accessories Charge Invoices"
    Else
        If COMPANY_CODE = "DJM" Or COMPANY_CODE = "HCA" Then
            pixServiceInvoices.Visible = False
            pixPartsInvoice.Visible = False
            pixSalesInvoices.Visible = False
            cmdOK.Value = True
        Else
            pixServiceInvoices.Visible = True
            pixPartsInvoice.Visible = False
            pixSalesInvoices.Visible = False
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrorCode:

    If cboSalesInvoiceType.Text = "Sales Invoice" Then
        If optVehicleSalesInvoices.Value = True Then
            If Module_Access(LOGID, "VEHICLE SALES INVOICE", "REPORTS") = False Then Exit Sub
            INVOICE_Type = "VEHICLE"
            FormExistsShow frmAMISSalesByInvoiceType
        Else
            If Module_Access(LOGID, "HYUNDAI VEHICLE SALES INVOICE", "REPORTS") = False Then Exit Sub
        End If
    ElseIf cboSalesInvoiceType.Text = "Parts Invoice" Then
        If optPartsCashInvoices.Value = True Then
            If Module_Access(LOGID, "PARTS CASH INVOICE", "REPORTS") = False Then Exit Sub
            INVOICE_Type = "PARTS-CASH"
            FormExistsShow frmAMISSalesByInvoiceType
        ElseIf optPartsChargeInvoices.Value = True Then
            If Module_Access(LOGID, "PARTS CHARGE INVOICE", "REPORTS") = False Then Exit Sub
            INVOICE_Type = "PARTS-CHARGE"
            FormExistsShow frmAMISSalesByInvoiceType
        ElseIf optHyundaiPartsCashInvoices.Value = True Then
            If Module_Access(LOGID, "HYUNDAI PARTS CASH INVOICE", "REPORTS") = False Then Exit Sub
        Else
            If Module_Access(LOGID, "HYUNDAI PARTS CHARGE INVOICE", "REPORTS") = False Then Exit Sub
        End If
    ElseIf cboSalesInvoiceType.Text = "Materials Invoice" Then
        If optPartsCashInvoices.Value = True Then
            'If Module_Access(LOGID, "MATERIALS CASH INVOICE", "REPORTS") = False Then Exit Sub
            INVOICE_Type = "MATERIALS-CASH"
            FormExistsShow frmAMISSalesByInvoiceType
        ElseIf optPartsChargeInvoices.Value = True Then
            'If Module_Access(LOGID, "MATERIALS CHARGE INVOICE", "REPORTS") = False Then Exit Sub
            INVOICE_Type = "MATERIALS-CHARGE"
            FormExistsShow frmAMISSalesByInvoiceType
        ElseIf optHyundaiPartsCashInvoices.Value = True Then
            If Module_Access(LOGID, "HYUNDAI MATERIALS CASH INVOICE", "REPORTS") = False Then Exit Sub
        Else
            If Module_Access(LOGID, "HYUNDAI MATERIALS CHARGE INVOICE", "REPORTS") = False Then Exit Sub
        End If
    ElseIf cboSalesInvoiceType.Text = "Accessories Invoice" Then
        If optPartsCashInvoices.Value = True Then
            'If Module_Access(LOGID, "ACCESSORIES CASH INVOICE", "REPORTS") = False Then Exit Sub
            INVOICE_Type = "ACCESSORIES-CASH"
            FormExistsShow frmAMISSalesByInvoiceType
        ElseIf optPartsChargeInvoices.Value = True Then
            'If Module_Access(LOGID, "ACCESSORIES CHARGE INVOICE", "REPORTS") = False Then Exit Sub
            INVOICE_Type = "ACCESSORIES-CHARGE"
            FormExistsShow frmAMISSalesByInvoiceType
        ElseIf optHyundaiPartsCashInvoices.Value = True Then
            If Module_Access(LOGID, "HYUNDAI ACCESSORIES CASH INVOICE", "REPORTS") = False Then Exit Sub
        Else
            If Module_Access(LOGID, "HYUNDAI ACCESSORIES CHARGE INVOICE", "REPORTS") = False Then Exit Sub
        End If
    ElseIf cboSalesInvoiceType.Text = "Service Invoice" Then
        If optServiceCashInvoices.Value = True Then
            If Module_Access(LOGID, "SERVICE CASH INVOICE", "REPORTS") = False Then Exit Sub
            INVOICE_Type = "SERVICE-CASH"
            FormExistsShow frmAMISSalesByInvoiceType
        ElseIf optServiceChargeInvoices.Value = True Then
            If Module_Access(LOGID, "SERVICE CHARGE INVOICE", "REPORTS") = False Then Exit Sub
            INVOICE_Type = "SERVICE-CHARGE"
            FormExistsShow frmAMISSalesByInvoiceType
        ElseIf optHyundaiServiceCashInvoices.Value = True Then
            If Module_Access(LOGID, "HYUNDAI SERVICE CASH INVOICE", "REPORTS") = False Then Exit Sub
        Else
            If Module_Access(LOGID, "HYUNDAI SERVICE CHARGE INVOICE", "REPORTS") = False Then Exit Sub
        End If
    End If

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    FillCombo
    cboSalesInvoiceType.Text = "Sales Invoice"
    pixPartsInvoice.Visible = False
    pixServiceInvoices.Visible = False
    CenterMe frmMain, Me, 1
End Sub


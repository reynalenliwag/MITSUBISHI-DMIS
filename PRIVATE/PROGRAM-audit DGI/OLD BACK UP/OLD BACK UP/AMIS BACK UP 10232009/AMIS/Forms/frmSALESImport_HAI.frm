VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Begin VB.Form frmSALESImport_HAI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALES Import Process"
   ClientHeight    =   7875
   ClientLeft      =   345
   ClientTop       =   1110
   ClientWidth     =   14040
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmSALESImport_HAI.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   14040
   Begin VB.CommandButton cmdClearJournals 
      BackColor       =   &H0080FF80&
      Caption         =   "Clear Selected Date"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11970
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdShowTrans 
      Caption         =   "Show Transactions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3660
      MouseIcon       =   "frmSALESImport_HAI.frx":030A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Process Import of SALES"
      Top             =   120
      Width           =   2010
   End
   Begin FlexCell.Grid Grid1 
      Height          =   4905
      Left            =   90
      TabIndex        =   9
      Top             =   1230
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8652
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontSize =   8.25
      ReadOnlyFocusRect=   0
      Rows            =   2
   End
   Begin wizProgBar.Prg progCPB 
      Height          =   315
      Left            =   8040
      TabIndex        =   2
      Top             =   6540
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   556
      Picture         =   "frmSALESImport_HAI.frx":045C
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "frmSALESImport_HAI.frx":0478
      ShowText        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpTranDate 
      Height          =   405
      Left            =   1860
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   51970049
      CurrentDate     =   38216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   13185
      MouseIcon       =   "frmSALESImport_HAI.frx":0494
      MousePointer    =   99  'Custom
      Picture         =   "frmSALESImport_HAI.frx":05E6
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Exit Window"
      Top             =   6945
      Width           =   720
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Import"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   12480
      MouseIcon       =   "frmSALESImport_HAI.frx":094C
      MousePointer    =   99  'Custom
      Picture         =   "frmSALESImport_HAI.frx":0A9E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Process Import of SALES"
      Top             =   6945
      Width           =   720
   End
   Begin wizProgBar.Prg Prg1 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   6540
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   556
      Picture         =   "frmSALESImport_HAI.frx":0D39
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarColor        =   4194304
      BarPicture      =   "frmSALESImport_HAI.frx":0D55
      ShowText        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
   Begin FlexCell.Grid Grid2 
      Height          =   4905
      Left            =   4740
      TabIndex        =   10
      Top             =   1230
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8652
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontSize =   8.25
      ReadOnlyFocusRect=   0
      Rows            =   2
   End
   Begin FlexCell.Grid Grid3 
      Height          =   4905
      Left            =   9390
      TabIndex        =   11
      Top             =   1230
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8652
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontSize =   8.25
      ReadOnlyFocusRect=   0
      Rows            =   2
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SMIS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   9390
      TabIndex        =   14
      Top             =   720
      Width           =   4545
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CSMS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   4740
      TabIndex        =   13
      Top             =   720
      Width           =   4545
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PMIS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   4515
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Only Un-Imported Invoices can be Imported"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   120
      TabIndex        =   8
      Top             =   7230
      Width           =   7995
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Over All Completion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   150
      TabIndex        =   7
      Top             =   6270
      Width           =   5835
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   210
      Width           =   1875
   End
   Begin VB.Label labCPB 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Completion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   8070
      TabIndex        =   1
      Top             =   6270
      Width           =   5835
   End
End
Attribute VB_Name = "frmSALESImport_HAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gconBIRData                                        As ADODB.Connection
Dim rsPMIOS_ORD_HD                                     As ADODB.Recordset
Dim rsTdayTran As ADODB.Recordset
Dim rsCSMIOS_REPOR As ADODB.Recordset
Dim rsSMIS_PURCHAGREE As ADODB.Recordset

Dim rsJournal_HDDup                                As ADODB.Recordset
Dim I                                              As Long
Dim GridImport As Integer
Dim TOTAL_DEBIT, TOTAL_CREDIT                      As Double
Dim J_JDATE As String, J_VOUCHERNO As String, J_JTYPE As String
Dim J_JNO As String, J_REMARKS As String, J_VENDORCODE As String, J_CUSTOMERCODE As String
Dim J_OUTBALANCE, J_AMOUNTTOPAY, J_INVOICEAMT, J_BALANCE, J_AMOUNTPAID As Double
Dim J_CHECKNO                                      As String
Dim J_INVOICEDATE As String, J_DUEDATE As String, J_PAYTYPE As String
Dim J_INVOICETYPE, J_INVOICENO                     As String
Dim J_CHECKDATE, J_BANKCODE                        As String
Dim J_REFNO, J_REFDATE                             As String
Dim J_TERMS, J_DEALER                              As String
Dim J_PAIDSTATUS, J_RECEIVESTATUS                  As String

Dim J_ACCT_CODE, J_ACCT_NAME                       As String
Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET       As Double
Dim J_STATUS, J_JITEMNO                            As String

Private Sub cmdCheck_Click()
    
    If Function_Access(LOGID, "Acess_Process", "IMPORT SALES ENTRIES") = False Then Exit Sub
    Screen.MousePointer = 11
    Prg1.Max = 3:
    Prg1.Value = 0: cmdCheck.Enabled = False: cmdExit.Enabled = False
    
    Dim WARRANTY_JNO As String
    Dim WARRANTY_VOUCHERNO  As String
    Dim WARRANTY_ItemCnt As Integer
    Dim WARRANTY_J_JITEMNO As String
    
    Dim WARRANTY_J_AMOUNTTOPAY As Double
    Dim WARRANTY_J_INVOICEAMT As Double
    Dim WARRANTY_J_BALANCE As Double
    Dim WARRANTY_J_AMOUNTPAID As Double
    
    Call ImportPMISSales
    Prg1.Value = 1
    
    Dim rsCSMIOS_REPOR                                 As ADODB.Recordset
    Dim rsCSMIOS_TINSPAINT                             As ADODB.Recordset
    Dim rsCSMIOS_SUBLET                                As ADODB.Recordset
    Dim rsCSMIOS_AIRCON                                As ADODB.Recordset
    Dim rsCSMIOS_LABOR                                 As ADODB.Recordset
    Dim rsCSMIOS_PARTS                                 As ADODB.Recordset
    Dim rsCSMIOS_MATERIALS                             As ADODB.Recordset
    Dim rsCSMIOS_ACCESSORIES                           As ADODB.Recordset

    Dim CSMIOS_REP_OR                                  As String
    Dim CSMIOS_ACCT_NO                                 As String
    Dim CSMIOS_PLATE_NO                                As String
    Dim CSMIOS_NIYM                                    As String
    Dim CSMIOS_PARTICIPAT                              As String
    Dim CSMIOS_TERM                                    As String
    Dim CSMIOS_DTE_REL                                 As String
    Dim CSMIOS_INVOICE                                 As String

    Dim CSMIOS_LABOR                                   As Double
    Dim CSMIOS_PARTS                                   As Double
    Dim CSMIOS_MATERIALS                               As Double
    Dim CSMIOS_ACCESSORIES                             As Double

    Dim CSMIOS_LABOR_COST                              As Double
    Dim CSMIOS_PARTS_COST                              As Double
    Dim CSMIOS_MATERIALS_COST                          As Double
    Dim CSMIOS_ACCESSORIES_COST                        As Double

    Dim CSMIOS_RO_AMOUNT                               As Double

    Dim CSMIOS_TINSPAINT                               As Double
    Dim CSMIOS_SUBLET                                  As Double
    Dim CSMIOS_AIRCON                                  As Double
    Dim CSMIOS_PMS                                     As Double

    'FOR PDI
    Dim CSMIOS_PDI_LABOR                               As Double
    Dim CSMIOS_PDI_PARTS                               As Double
    Dim CSMIOS_PDI_MATERIALS                           As Double

    Dim CSMIOS_PDI_TINSPAINT                           As Double
    Dim CSMIOS_PDI_SUBLET                              As Double
    Dim CSMIOS_PDI_AIRCON                              As Double
    'END PDI
    
    Dim CSMIOS_TINSPAINT_DISCOUNT                      As Double
    Dim CSMIOS_SUBLET_DISCOUNT                         As Double
    Dim CSMIOS_AIRCON_DISCOUNT                         As Double

    Dim CSMIOS_LABOR_DISCOUNT                          As Double
    Dim CSMIOS_PARTS_DISCOUNT                          As Double
    Dim CSMIOS_MATERIALS_DISCOUNT                      As Double
    Dim CSMIOS_ACCESSORIES_DISCOUNT                    As Double

    Dim WARRANTY_DIRECT_EXPENSE_LABOR                  As Double
    Dim WARRANTY_DIRECT_EXPENSE_SPAREPARTS             As Double
    Dim WARRANTY_DIRECT_EXPENSE_GOL                    As Double
    Dim WARRANTY_DIRECT_EXPENSE_ACCESSORIES            As Double
    
    Dim WARRANTY_CSMIOS_PARTS_COST                     As Double
    Dim WARRANTY_CSMIOS_MATERIALS_COST                 As Double
    Dim WARRANTY_CSMIOS_ACCESSORIES_COST               As Double
                
    Dim COMPANY_DIRECT_EXPENSE_LABOR                   As Double
    Dim COMPANY_DIRECT_EXPENSE_SPAREPARTS              As Double
    Dim COMPANY_DIRECT_EXPENSE_GOL                     As Double
    Dim COMPANY_DIRECT_EXPENSE_ACCESSORIES             As Double

    Dim SALES_DIRECT_EXPENSE_LABOR                     As Double
    Dim SALES_DIRECT_EXPENSE_SPAREPARTS                As Double
    Dim SALES_DIRECT_EXPENSE_GOL                       As Double
    Dim SALES_DIRECT_EXPENSE_ACCESSORIES               As Double

    Dim INSURANCE_DIRECT_EXPENSE_LABOR                 As Double
    Dim INSURANCE_DIRECT_EXPENSE_SPAREPARTS            As Double
    Dim INSURANCE_DIRECT_EXPENSE_GOL                   As Double
    Dim INSURANCE_DIRECT_EXPENSE_ACCESSORIES           As Double
    Dim TOTAL_INSURANCE_AMOUNT                         As Double
    
    Dim CSMS_VAT_EXEMPT                                As Boolean
    Dim ItemCnt                                        As Integer
    
    I = 0
    For GridImport = 1 To Grid2.Rows - 1
        If N2Str2Zero(Grid2.Cell(GridImport, 1).Text) = 0 Then
            Set rsCSMIOS_REPOR = New ADODB.Recordset
            Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where invoice = '" & Grid2.Cell(GridImport, 3).Text & "' and dte_comp = '" & CDate(dtpTranDate) & "' order by invoice ASC")
            If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
                ItemCnt = 0
                CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
                CSMIOS_ACCT_NO = Null2String(rsCSMIOS_REPOR!ACCT_NO)
                
                CSMIOS_PARTICIPAT = Null2String(rsCSMIOS_REPOR!PARTICIPAT)
    
                CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!Plate_No)
                CSMIOS_NIYM = Null2String(rsCSMIOS_REPOR!Niym)
                CSMIOS_TERM = Null2String(rsCSMIOS_REPOR!TERM)
                CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!DTE_COMP)
                CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!Invoice)
                
                CSMS_VAT_EXEMPT = Null2Bool(rsCSMIOS_REPOR!VAT_EXEMPT)
                
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If
    
                CSMIOS_RO_AMOUNT = Round(N2Str2Zero(rsCSMIOS_REPOR!RO_AMOUNT), 2)
            
                
                '=======================================================================================================================================================================
                'CUSTOMER
                
                'LABOR - MECHANICAL / BODY AND PAINT / AIRCON
                Set rsCSMIOS_LABOR = New ADODB.Recordset
                Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABOR Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then
                    CSMIOS_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)
                    CSMIOS_LABOR_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_LABOR!discount), 2)
                Else
                    CSMIOS_LABOR = 0: CSMIOS_LABOR_DISCOUNT = 0
                End If
                
                Set rsCSMIOS_TINSPAINT = New ADODB.Recordset
                Set rsCSMIOS_TINSPAINT = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS TINSPAINT,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_TINSPAINT Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_TINSPAINT.EOF And Not rsCSMIOS_TINSPAINT.BOF Then
                    CSMIOS_TINSPAINT = Round(N2Str2Zero(rsCSMIOS_TINSPAINT!TINSPAINT), 2)
                    CSMIOS_TINSPAINT_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_TINSPAINT!discount), 2)
                Else
                    CSMIOS_TINSPAINT = 0: CSMIOS_TINSPAINT_DISCOUNT = 0
                End If
    
                Set rsCSMIOS_AIRCON = New ADODB.Recordset
                Set rsCSMIOS_AIRCON = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS AIRCON,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_AIRCON Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_AIRCON.EOF And Not rsCSMIOS_AIRCON.BOF Then
                    CSMIOS_AIRCON = Round(N2Str2Zero(rsCSMIOS_AIRCON!AIRCON), 2)
                    CSMIOS_AIRCON_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_AIRCON!discount), 2)
                Else
                    CSMIOS_AIRCON = 0: CSMIOS_AIRCON_DISCOUNT = 0
                End If
        
                'PARTS
                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTS Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                    CSMIOS_PARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)
                    CSMIOS_PARTS_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_PARTS!discount), 2)
                Else
                    CSMIOS_PARTS = 0: CSMIOS_PARTS_DISCOUNT = 0
                End If
    
                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS PARTS_COST from CSMIOS_vw_PARTSCOST Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                    CSMIOS_PARTS_COST = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS_COST), 2)
                Else
                    CSMIOS_PARTS_COST = 0:
                End If
    
                'MATERIALS
                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALS Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
                    CSMIOS_MATERIALS = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)
                    CSMIOS_MATERIALS_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_MATERIALS!discount), 2)
                Else
                    CSMIOS_MATERIALS = 0: CSMIOS_MATERIALS_DISCOUNT = 0
                End If
                
                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS MAT_COST from CSMIOS_vw_MATCOST Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                    CSMIOS_MATERIALS_COST = Round(N2Str2Zero(rsCSMIOS_PARTS!MAT_COST), 2)
                Else
                    CSMIOS_MATERIALS_COST = 0:
                End If
    
                'ACCESSORIES
                Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
                Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS ACCESSORIES,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_ACCESSORIES Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                    CSMIOS_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACCESSORIES), 2)
                    CSMIOS_ACCESSORIES_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!discount), 2)
                Else
                    CSMIOS_ACCESSORIES = 0: CSMIOS_ACCESSORIES_DISCOUNT = 0
                End If
                
                Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
                Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS ACC_COST from CSMIOS_vw_ACCCOST Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                    CSMIOS_ACCESSORIES_COST = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACC_COST), 2)
                Else
                    CSMIOS_ACCESSORIES_COST = 0:
                End If
    
                '=======================================================================================================================================================================
                'WARRANTY
                
                WARRANTY_DIRECT_EXPENSE_LABOR = 0: WARRANTY_DIRECT_EXPENSE_SPAREPARTS = 0: WARRANTY_DIRECT_EXPENSE_GOL = 0: WARRANTY_DIRECT_EXPENSE_ACCESSORIES = 0
                WARRANTY_CSMIOS_PARTS_COST = 0: WARRANTY_CSMIOS_MATERIALS_COST = 0: WARRANTY_CSMIOS_ACCESSORIES_COST = 0
                
                'LABOR
                Set rsCSMIOS_LABOR = New ADODB.Recordset
                Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORWarranty Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then
                   WARRANTY_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)
                End If
    
                'PARTS
                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTSWarranty Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                   WARRANTY_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)
                End If
                
                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS PARTS_COST from CSMIOS_vw_WARRANTY_PARTSCOST Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                    WARRANTY_CSMIOS_PARTS_COST = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS_COST), 2)
                Else
                    WARRANTY_CSMIOS_PARTS_COST = 0:
                End If
                
                'MATERIALS
                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALSWarranty Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
                   WARRANTY_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)
                End If
        
                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS MAT_COST from CSMIOS_vw_WARRANTY_MATCOST Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
                    WARRANTY_CSMIOS_MATERIALS_COST = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MAT_COST), 2)
                Else
                    WARRANTY_CSMIOS_MATERIALS_COST = 0:
                End If
    
                'ACCESSORIES
                
                Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
                Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS ACCESSORIES,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_ACCESSORIESWarranty Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                   WARRANTY_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACCESSORIES), 2)
                End If
                
                Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
                Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS ACC_COST from CSMIOS_vw_WARRANTY_ACCCOST Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                    WARRANTY_CSMIOS_ACCESSORIES_COST = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACC_COST), 2)
                Else
                    WARRANTY_CSMIOS_ACCESSORIES_COST = 0:
                End If
        
                '=======================================================================================================================================================================
                
                '=======================================================================================================================================================================
                'COMPANY
                
                COMPANY_DIRECT_EXPENSE_LABOR = 0: COMPANY_DIRECT_EXPENSE_SPAREPARTS = 0: COMPANY_DIRECT_EXPENSE_GOL = 0: COMPANY_DIRECT_EXPENSE_ACCESSORIES = 0
                
                Set rsCSMIOS_LABOR = New ADODB.Recordset
                Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then
                   COMPANY_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)
                End If
                
                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTSCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                   COMPANY_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)
                End If
                
                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALSCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
                   COMPANY_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)
                End If
                
                Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
                Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS ACCESSORIES,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_ACCESSORIESCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                   COMPANY_DIRECT_EXPENSE_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACCESSORIES), 2)
                End If
                
                '=======================================================================================================================================================================
    
                '=======================================================================================================================================================================
                'SALES
                
                SALES_DIRECT_EXPENSE_LABOR = 0: SALES_DIRECT_EXPENSE_SPAREPARTS = 0: SALES_DIRECT_EXPENSE_GOL = 0: SALES_DIRECT_EXPENSE_ACCESSORIES = 0
    
                Set rsCSMIOS_LABOR = New ADODB.Recordset
                Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then
                   SALES_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)
                End If
    
                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTSSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                   SALES_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)
                End If
                
                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALSSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
                   SALES_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)
                End If
    
                Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
                Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS ACCESSORIES,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_ACCESSORIESSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                   SALES_DIRECT_EXPENSE_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACCESSORIES), 2)
                End If
                
                INSURANCE_DIRECT_EXPENSE_LABOR = 0: INSURANCE_DIRECT_EXPENSE_SPAREPARTS = 0: INSURANCE_DIRECT_EXPENSE_GOL = 0: INSURANCE_DIRECT_EXPENSE_ACCESSORIES = 0
                TOTAL_INSURANCE_AMOUNT = 0
                
                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select * from CSMIOS_INSURANCE Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
                   INSURANCE_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_MATERIALS!INSLABOR), 2)
                   INSURANCE_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!INSMATERIALS), 2)
                   INSURANCE_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_MATERIALS!INSPARTS), 2)
                   INSURANCE_DIRECT_EXPENSE_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_MATERIALS!INSACCESSORIES), 2)
                   
                   If (CSMIOS_LABOR + CSMIOS_SUBLET + CSMIOS_TINSPAINT + CSMIOS_PMS) - INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                        If CSMIOS_LABOR - INSURANCE_DIRECT_EXPENSE_LABOR >= 0 Then
                           CSMIOS_LABOR = Round(CSMIOS_LABOR - INSURANCE_DIRECT_EXPENSE_LABOR, 2)
                           GoTo PAKSIW
                        Else
                           If CSMIOS_LABOR - INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                              CSMIOS_LABOR = Round(CSMIOS_LABOR - INSURANCE_DIRECT_EXPENSE_LABOR, 2)
                              GoTo PAKSIW
                           Else
                              INSURANCE_DIRECT_EXPENSE_LABOR = Round(INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_LABOR, 2)
                              CSMIOS_LABOR = 0
                           End If
                        End If
                        If CSMIOS_SUBLET > 0 And CSMIOS_LABOR - INSURANCE_DIRECT_EXPENSE_LABOR >= 0 Then
                           CSMIOS_SUBLET = Round(CSMIOS_SUBLET - (INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_LABOR), 2)
                           GoTo PAKSIW
                        Else
                           If CSMIOS_SUBLET - INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                              CSMIOS_SUBLET = Round(CSMIOS_SUBLET - INSURANCE_DIRECT_EXPENSE_LABOR, 2)
                              GoTo PAKSIW
                           Else
                              INSURANCE_DIRECT_EXPENSE_LABOR = Round(INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_SUBLET, 2)
                              CSMIOS_SUBLET = 0
                           End If
                        End If
                        If CSMIOS_TINSPAINT > 0 And CSMIOS_LABOR - CSMIOS_SUBLET - INSURANCE_DIRECT_EXPENSE_LABOR >= 0 Then
                           CSMIOS_TINSPAINT = Round(CSMIOS_TINSPAINT - (INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_LABOR - CSMIOS_SUBLET), 2)
                           GoTo PAKSIW
                        Else
                           If CSMIOS_TINSPAINT - INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                              CSMIOS_TINSPAINT = Round(CSMIOS_TINSPAINT - INSURANCE_DIRECT_EXPENSE_LABOR, 2)
                              GoTo PAKSIW
                           Else
                              INSURANCE_DIRECT_EXPENSE_LABOR = Round(INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_TINSPAINT, 2)
                              CSMIOS_TINSPAINT = 0
                           End If
                        End If
                        If CSMIOS_PMS > 0 And CSMIOS_LABOR - CSMIOS_SUBLET - CSMIOS_TINSPAINT - INSURANCE_DIRECT_EXPENSE_LABOR >= 0 Then
                           CSMIOS_PMS = Round(CSMIOS_PMS - (INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_LABOR - CSMIOS_SUBLET - CSMIOS_TINSPAINT), 2)
                           GoTo PAKSIW
                        Else
                           If CSMIOS_PMS - INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                              CSMIOS_PMS = Round(CSMIOS_PMS - INSURANCE_DIRECT_EXPENSE_LABOR, 2)
                           Else
                              INSURANCE_DIRECT_EXPENSE_LABOR = Round(INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_PMS, 2)
                              CSMIOS_PMS = 0
                           End If
                        End If
PAKSIW:                 INSURANCE_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_MATERIALS!INSLABOR), 2)
                   Else
                       CSMIOS_LABOR = 0
                       CSMIOS_SUBLET = 0
                       CSMIOS_TINSPAINT = 0
                       CSMIOS_PMS = 0
                   End If
                   If CSMIOS_PARTS > 0 Then
                      CSMIOS_PARTS = Round(CSMIOS_PARTS - INSURANCE_DIRECT_EXPENSE_SPAREPARTS, 2)
                   End If
                   If CSMIOS_MATERIALS > 0 Then
                      CSMIOS_MATERIALS = Round(CSMIOS_MATERIALS - INSURANCE_DIRECT_EXPENSE_GOL, 2)
                   End If
                   If CSMIOS_ACCESSORIES > 0 Then
                      CSMIOS_ACCESSORIES = Round(CSMIOS_ACCESSORIES - INSURANCE_DIRECT_EXPENSE_ACCESSORIES, 2)
                   End If
                   
                   TOTAL_INSURANCE_AMOUNT = Round(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES, 2)
                End If
                '====================================================================================================================================================================================
                
                '=======================================================================================================================================================================
                
                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
    
                'HEADER
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                
                J_JDATE = N2Date2Null(CSMIOS_DTE_REL)
                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                J_JTYPE = "'SJ'"
                J_REMARKS = "NULL"
                J_VENDORCODE = "'999999'"
                J_CUSTOMERCODE = N2Str2Null(CSMIOS_ACCT_NO)
    
                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0
    
                J_AMOUNTTOPAY = 0
                J_INVOICEAMT = Round(NumericVal(CSMIOS_RO_AMOUNT - TOTAL_INSURANCE_AMOUNT), 2)
                J_BALANCE = Round(NumericVal(CSMIOS_RO_AMOUNT - TOTAL_INSURANCE_AMOUNT), 2)
                J_AMOUNTPAID = 0
    
                J_STATUS = "'N'"
    
                J_INVOICEDATE = N2Date2Null(CSMIOS_DTE_REL)
                J_INVOICENO = N2Str2Null(CSMIOS_INVOICE)
                J_CHECKNO = "NULL"
                J_DUEDATE = N2Date2Null(CSMIOS_DTE_REL)
                J_PAYTYPE = "NULL"
                J_INVOICETYPE = "'SI'"
                J_CHECKDATE = "NULL"
                J_BANKCODE = "NULL"
                J_REFNO = N2Str2Null(CSMIOS_REP_OR)
                J_REFDATE = N2Date2Null(CSMIOS_DTE_REL)
                J_TERMS = N2Str2Null(CSMIOS_TERM)
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"
    
                If J_INVOICEAMT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_ACCT_CODE = N2Str2Null(COA_PRE_DELIVERY)
                        J_ACCT_NAME = N2Str2Null(Setacctname(COA_PRE_DELIVERY))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                    End If
                    J_DEBIT = Round(J_INVOICEAMT, 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_LABOR > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR")))
                    End If
                    J_GROSS = NumericVal(CSMIOS_LABOR)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_LABOR), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_LABOR), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_LABOR / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_LABOR) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_LABOR_DISCOUNT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "LABOR"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "LABOR")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "LABOR"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "LABOR")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_LABOR_DISCOUNT), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_LABOR_DISCOUNT), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_LABOR_DISCOUNT), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_LABOR_DISCOUNT / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_LABOR_DISCOUNT) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_TINSPAINT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "BODY"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "BODY")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "BODY"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "BODY")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_TINSPAINT), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_TINSPAINT), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_TINSPAINT), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_TINSPAINT / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_TINSPAINT) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_TINSPAINT_DISCOUNT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "BODY"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "BODY")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "BODY"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "BODY")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_TINSPAINT_DISCOUNT), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_TINSPAINT_DISCOUNT), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_TINSPAINT_DISCOUNT), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_TINSPAINT_DISCOUNT / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_TINSPAINT_DISCOUNT) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                     " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_SUBLET > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "SUBLET"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "SUBLET")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "SUBLET"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "SUBLET")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_SUBLET), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_SUBLET), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_SUBLET), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_SUBLET / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_SUBLET) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                End If
                If CSMIOS_SUBLET_DISCOUNT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "SUBLET"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "SUBLET")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "SUBLET"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "SUBLET")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_SUBLET_DISCOUNT), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_SUBLET_DISCOUNT), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_SUBLET_DISCOUNT), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_SUBLET_DISCOUNT / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_SUBLET_DISCOUNT) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_AIRCON > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "AIRCON"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "AIRCON")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "AIRCON"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "AIRCON")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_AIRCON), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_AIRCON), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_AIRCON), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_AIRCON / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_AIRCON) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_AIRCON_DISCOUNT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "AIRCON"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "AIRCON")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "AIRCON"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "AIRCON")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_AIRCON_DISCOUNT), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_AIRCON_DISCOUNT), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_AIRCON_DISCOUNT), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_AIRCON_DISCOUNT / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_AIRCON_DISCOUNT) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                               
                If CSMIOS_PARTS > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS")))
                    End If
                    J_GROSS = NumericVal(CSMIOS_PARTS)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_PARTS), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_PARTS), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_PARTS / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_PARTS) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "PARTS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "PARTS")))
                    End If
                    J_DEBIT = Round(CSMIOS_PARTS_COST, 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(CSMIOS_PARTS_COST, 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                
                If CSMIOS_PARTS_DISCOUNT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "PARTS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "PARTS")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_PARTS_DISCOUNT), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_PARTS_DISCOUNT), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_PARTS_DISCOUNT), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_PARTS_DISCOUNT / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_PARTS_DISCOUNT) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_MATERIALS > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_MATERIALS), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_MATERIALS), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_MATERIALS), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_MATERIALS / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_MATERIALS) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                     
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "LUBRICANTS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "LUBRICANTS")))
                    End If
                    J_GROSS = 0
                    J_TAX = 0
                    J_NET = 0
                    J_DEBIT = Round(CSMIOS_MATERIALS_COST, 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("LUBRICANTS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("LUBRICANTS")))
                    End If
                    J_GROSS = 0
                    J_TAX = 0
                    J_NET = 0
                    J_DEBIT = 0
                    J_CREDIT = Round(CSMIOS_MATERIALS_COST, 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_MATERIALS_DISCOUNT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_MATERIALS_DISCOUNT), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_MATERIALS_DISCOUNT), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_MATERIALS_DISCOUNT), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_MATERIALS_DISCOUNT / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_MATERIALS_DISCOUNT) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
    
                'ACCESSORIES
                If CSMIOS_ACCESSORIES > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "ACCESSORIES")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "ACCESSORIES")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_ACCESSORIES), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_ACCESSORIES), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_ACCESSORIES), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_ACCESSORIES / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_ACCESSORIES) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                     
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("PARTS", "ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("PARTS", "ACCESSORIES")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("PARTS", "ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("PARTS", "ACCESSORIES")))
                    End If
                    J_GROSS = 0
                    J_TAX = 0
                    J_NET = 0
                    J_DEBIT = Round(CSMIOS_ACCESSORIES_COST, 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES")))
                    End If
                    J_GROSS = 0
                    J_TAX = 0
                    J_NET = 0
                    J_DEBIT = 0
                    J_CREDIT = Round(CSMIOS_ACCESSORIES_COST, 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
                If CSMIOS_ACCESSORIES_DISCOUNT > 0 Then
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "ACCESSORIES")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "ACCESSORIES")))
                    End If
                    J_GROSS = Round(NumericVal(CSMIOS_ACCESSORIES_DISCOUNT), 2)
                    If CSMIOS_INVOICE = "PDI RO" Then
                        J_TAX = 0
                        J_NET = Round(NumericVal(CSMIOS_ACCESSORIES_DISCOUNT), 2)
                    Else
                        If CSMS_VAT_EXEMPT = True Then
                           J_TAX = 0
                           J_NET = Round(NumericVal(CSMIOS_ACCESSORIES_DISCOUNT), 2)
                        Else
                           J_TAX = Round(NumericVal(Round((CSMIOS_ACCESSORIES_DISCOUNT / 1.12), 2) * 0.12), 2)
                           J_NET = Round(NumericVal(CSMIOS_ACCESSORIES_DISCOUNT) - NumericVal(J_TAX), 2)
                        End If
                    End If
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                End If
    
                If J_INVOICEAMT > 0 Then
                    If CSMS_VAT_EXEMPT = False Then
                       ItemCnt = ItemCnt + 1
                       J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnOutPutTax())
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutPutTax()))
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(Round((J_INVOICEAMT / 1.12), 2) * 0.12), 2)
                       TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                       If Round(Round(TOTAL_CREDIT, 2) - Round(TOTAL_DEBIT, 2), 2) = 0.01 Then J_CREDIT = Round(J_CREDIT, 2) - 0.01
                       If Round(Round(TOTAL_DEBIT, 2) - Round(TOTAL_CREDIT, 2), 2) = 0.01 Then J_CREDIT = Round(J_CREDIT, 2) + 0.01
                       J_TAX = 0
                       J_GROSS = 0
                       J_NET = 0
                       J_STATUS = "'N'"
        
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    End If
                    
                    gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                                     " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                     " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                     ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                
                End If
                
                '==================================================================================================================================================================================
                'WARRANTY
                If WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL + WARRANTY_DIRECT_EXPENSE_ACCESSORIES > 0 Then
                    TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
                    Set rsJournal_HDDup = New ADODB.Recordset
                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 2, "000000") & "'"
                    
                    WARRANTY_VOUCHERNO = N2Str2Null(Format(NumericVal(GetVoucherNo()) + 1, "000000"))
                    WARRANTY_ItemCnt = 0
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("WARRANTY"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("WARRANTY")))
                    J_DEBIT = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL + WARRANTY_DIRECT_EXPENSE_ACCESSORIES), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                     " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                    If WARRANTY_DIRECT_EXPENSE_LABOR > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR"))
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR")))
                       J_GROSS = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR), 2)
                       J_TAX = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR) / 9.3333, 2)
                       J_NET = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR) - NumericVal(J_TAX), 2)
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(J_NET), 2)
                       J_STATUS = "'N'"
                       TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    End If
                    If WARRANTY_DIRECT_EXPENSE_SPAREPARTS > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS"))
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS")))
                       J_GROSS = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_SPAREPARTS), 2)
                       J_TAX = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_SPAREPARTS) / 9.3333, 2)
                       J_NET = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_SPAREPARTS) - NumericVal(J_TAX), 2)
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(J_NET), 2)
                       J_STATUS = "'N'"
                       TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        If CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "PARTS")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "PARTS")))
                        End If
                        J_DEBIT = Round(WARRANTY_CSMIOS_PARTS_COST, 2)
                        J_CREDIT = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                       " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                         ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        If CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                        End If
                        J_DEBIT = 0
                        J_CREDIT = Round(WARRANTY_CSMIOS_PARTS_COST, 2)
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                       " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                         ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                    End If
                    If WARRANTY_DIRECT_EXPENSE_GOL > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                       J_GROSS = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_GOL), 2)
                       J_TAX = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_GOL) / 9.3333, 2)
                       J_NET = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_GOL) - NumericVal(J_TAX), 2)
                       J_DEBIT = 0
                       J_CREDIT = NumericVal(J_NET)
                       J_STATUS = "'N'"
                       TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        If CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "LUBRICANTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "LUBRICANTS")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "LUBRICANTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "LUBRICANTS")))
                        End If
                        J_GROSS = 0
                        J_TAX = 0
                        J_NET = 0
                        J_DEBIT = Round(WARRANTY_CSMIOS_MATERIALS_COST, 2)
                        J_CREDIT = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                       " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                         ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        If CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("LUBRICANTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("LUBRICANTS")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("LUBRICANTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("LUBRICANTS")))
                        End If
                        J_GROSS = 0
                        J_TAX = 0
                        J_NET = 0
                        J_DEBIT = 0
                        J_CREDIT = Round(WARRANTY_CSMIOS_MATERIALS_COST, 2)
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                       " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                         ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        
                    End If
    
                    If WARRANTY_DIRECT_EXPENSE_ACCESSORIES > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "ACCESSORIES"))
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "ACCESSORIES")))
                       J_GROSS = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_ACCESSORIES), 2)
                       J_TAX = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_ACCESSORIES) / 9.3333, 2)
                       J_NET = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_ACCESSORIES) - NumericVal(J_TAX), 2)
                       J_DEBIT = 0
                       J_CREDIT = NumericVal(J_NET)
                       J_STATUS = "'N'"
                       TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        If CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("PARTS", "ACCESSORIES"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("PARTS", "ACCESSORIES")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("PARTS", "ACCESSORIES"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("PARTS", "ACCESSORIES")))
                        End If
                        J_GROSS = 0
                        J_TAX = 0
                        J_NET = 0
                        J_DEBIT = Round(WARRANTY_CSMIOS_ACCESSORIES_COST, 2)
                        J_CREDIT = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                       " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                         ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        If CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES")))
                        End If
                        J_GROSS = 0
                        J_TAX = 0
                        J_NET = 0
                        J_DEBIT = 0
                        J_CREDIT = Round(WARRANTY_CSMIOS_ACCESSORIES_COST, 2)
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                       " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                         ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        
                    End If
    
                    If NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL + WARRANTY_DIRECT_EXPENSE_ACCESSORIES) > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = COA_OUTPUT_TAX
                       J_ACCT_NAME = N2Str2Null(Setacctname(COA_OUTPUT_TAX))
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL + WARRANTY_DIRECT_EXPENSE_ACCESSORIES) / 9.3333), 2)
                       J_TAX = 0
                       J_GROSS = 0
                       J_NET = 0
                       J_STATUS = "'N'"
                       TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                       If Round(Round(TOTAL_CREDIT, 2) - Round(TOTAL_DEBIT, 2), 2) = 0.01 Then J_CREDIT = Round(J_CREDIT, 2) - 0.01
                       If Round(Round(TOTAL_DEBIT, 2) - Round(TOTAL_CREDIT, 2), 2) = 0.01 Then J_CREDIT = Round(J_CREDIT, 2) + 0.01
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    End If
                    
                    CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
                    CSMIOS_ACCT_NO = Null2String("H00001")
                    
                    CSMIOS_PARTICIPAT = ""
        
                    CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!Plate_No)
                    CSMIOS_NIYM = Null2String(SetCustomerName("H00001"))
                    CSMIOS_TERM = Null2String(rsCSMIOS_REPOR!TERM)
                    CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!DTE_COMP)
                    CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!Invoice)
                    
                    Set rsJournal_HDDup = New ADODB.Recordset
                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                        J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                    Else
                        J_JNO = "'000001'"
                    End If
        
                    CSMIOS_RO_AMOUNT = N2Str2Zero(rsCSMIOS_REPOR!RO_AMOUNT)
                    
                    J_JDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_JTYPE = "'SJ'"
                    J_REMARKS = "NULL"
                    J_VENDORCODE = "'999999'"
                    J_CUSTOMERCODE = N2Str2Null("H00001")
        
                    J_DEBIT = 0
                    J_CREDIT = 0
                    J_TAX = 0
                    J_OUTBALANCE = 0
        
                    J_AMOUNTTOPAY = 0
                    J_INVOICEAMT = Round(CSMIOS_RO_AMOUNT, 2)
                    J_BALANCE = Round(CSMIOS_RO_AMOUNT, 2)
                    J_AMOUNTPAID = 0
        
                    J_STATUS = "'N'"
        
                    J_INVOICEDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_INVOICENO = N2Str2Null(CSMIOS_INVOICE)
                    J_CHECKNO = "NULL"
                    J_DUEDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_PAYTYPE = "NULL"
                    J_INVOICETYPE = "'SI'"
                    J_CHECKDATE = "NULL"
                    J_BANKCODE = "NULL"
                    J_REFNO = N2Str2Null(CSMIOS_REP_OR)
                    J_REFDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_TERMS = N2Str2Null(CSMIOS_TERM)
                    J_DEALER = "NULL"
                    J_PAIDSTATUS = "'N'"
                    J_RECEIVESTATUS = "'N'"
                    
                    WARRANTY_J_AMOUNTTOPAY = 0
                    WARRANTY_J_INVOICEAMT = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL + WARRANTY_DIRECT_EXPENSE_ACCESSORIES), 2)
                    WARRANTY_J_BALANCE = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL + WARRANTY_DIRECT_EXPENSE_ACCESSORIES), 2)
                    WARRANTY_J_AMOUNTPAID = 0
                    gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                                     " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                     " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & WARRANTY_J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & Round(WARRANTY_J_AMOUNTTOPAY, 2) & "," & WARRANTY_J_BALANCE & "," & WARRANTY_J_AMOUNTPAID & _
                                     ", " & WARRANTY_JNO & ", " & Round(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL + WARRANTY_DIRECT_EXPENSE_ACCESSORIES, 2) & ", " & Round(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL + WARRANTY_DIRECT_EXPENSE_ACCESSORIES, 2) & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                End If
                '==================================================================================================================================================================================
                        
                '==================================================================================================================================================================================
                'INSURANCE
                
                If INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES > 0 Then
    
                    TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
                    Set rsJournal_HDDup = New ADODB.Recordset
                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 2, "000000") & "'"
                    
                    WARRANTY_VOUCHERNO = N2Str2Null(Format(NumericVal(GetVoucherNo()) + 1, "000000"))
                    WARRANTY_ItemCnt = 0
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("INSURANCE"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("INSURANCE")))
                    J_DEBIT = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                     " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                    If INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR"))
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR")))
                       J_GROSS = NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR)
                       J_TAX = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR) / 9.3333, 2)
                       J_NET = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR) - NumericVal(J_TAX), 2)
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(J_NET), 2)
                       J_STATUS = "'N'"
                       TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    End If
                    If INSURANCE_DIRECT_EXPENSE_SPAREPARTS > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS"))
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS")))
                       J_GROSS = NumericVal(INSURANCE_DIRECT_EXPENSE_SPAREPARTS)
                       J_TAX = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_SPAREPARTS) / 9.3333, 2)
                       J_NET = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_SPAREPARTS) - NumericVal(J_TAX), 2)
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(J_NET), 2)
                       J_STATUS = "'N'"
                       TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        
                    
                    End If
                    If INSURANCE_DIRECT_EXPENSE_GOL > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                       J_GROSS = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_GOL), 2)
                       J_TAX = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_GOL) / 9.3333, 2)
                       J_NET = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_GOL) - NumericVal(J_TAX), 2)
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(J_NET), 2)
                       J_STATUS = "'N'"
                       TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                        
                    End If
    
                    If INSURANCE_DIRECT_EXPENSE_ACCESSORIES > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "ACCESSORIES"))
                       J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "ACCESSORIES")))
                       J_GROSS = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_ACCESSORIES), 2)
                       J_TAX = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_ACCESSORIES) / 9.3333, 2)
                       J_NET = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_ACCESSORIES) - NumericVal(J_TAX), 2)
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(J_NET), 2)
                       J_STATUS = "'N'"
                       TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    
                        
                    End If
    
                    If NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES) > 0 Then
                       WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                       WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                       J_ACCT_CODE = COA_OUTPUT_TAX
                       J_ACCT_NAME = N2Str2Null(Setacctname(COA_OUTPUT_TAX))
                       J_DEBIT = 0
                       J_CREDIT = Round(NumericVal(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES) / 9.3333), 2)
                       J_TAX = 0
                       J_GROSS = 0
                       J_NET = 0
                       J_STATUS = "'N'"
                       TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                       If Round(Round(TOTAL_CREDIT, 2) - Round(TOTAL_DEBIT, 2), 2) = 0.01 Then J_CREDIT = Round(J_CREDIT, 2) - 0.01
                       If Round(Round(TOTAL_DEBIT, 2) - Round(TOTAL_CREDIT, 2), 2) = 0.01 Then J_CREDIT = Round(J_CREDIT, 2) + 0.01
                         
                       gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    End If
                    
                    CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
                    CSMIOS_ACCT_NO = Null2String("H00001")
                        
                    CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!Plate_No)
                    CSMIOS_NIYM = Null2String(SetCustomerName(CSMIOS_PARTICIPAT))
                    CSMIOS_TERM = Null2String(rsCSMIOS_REPOR!TERM)
                    CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!DTE_COMP)
                    CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!Invoice)
                    
                    Set rsJournal_HDDup = New ADODB.Recordset
                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                        J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                    Else
                        J_JNO = "'000001'"
                    End If
        
                    CSMIOS_RO_AMOUNT = N2Str2Zero(rsCSMIOS_REPOR!RO_AMOUNT)
                    
                    J_JDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                    J_JTYPE = "'SJ'"
                    J_REMARKS = "NULL"
                    J_VENDORCODE = "'999999'"
        
                    J_DEBIT = 0
                    J_CREDIT = 0
                    J_TAX = 0
                    J_OUTBALANCE = 0
        
                    J_AMOUNTTOPAY = 0
                    J_INVOICEAMT = CSMIOS_RO_AMOUNT
                    J_BALANCE = CSMIOS_RO_AMOUNT
                    J_AMOUNTPAID = 0
        
                    J_STATUS = "'N'"
        
                    J_INVOICEDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_INVOICENO = N2Str2Null(CSMIOS_INVOICE)
                    J_CHECKNO = "NULL"
                    J_DUEDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_PAYTYPE = "NULL"
                    J_INVOICETYPE = "'SI'"
                    J_CHECKDATE = "NULL"
                    J_BANKCODE = "NULL"
                    J_REFNO = N2Str2Null(CSMIOS_REP_OR)
                    J_REFDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_TERMS = N2Str2Null(CSMIOS_TERM)
                    J_DEALER = "NULL"
                    J_PAIDSTATUS = "'N'"
                    J_RECEIVESTATUS = "'N'"
                    
                    WARRANTY_J_AMOUNTTOPAY = 0
                    WARRANTY_J_INVOICEAMT = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES), 2)
                    WARRANTY_J_BALANCE = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES), 2)
                    WARRANTY_J_AMOUNTPAID = 0
                    gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                                     " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                     " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & N2Str2Null(CSMIOS_PARTICIPAT) & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & WARRANTY_J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & WARRANTY_J_AMOUNTTOPAY & "," & WARRANTY_J_BALANCE & "," & WARRANTY_J_AMOUNTPAID & _
                                     ", " & WARRANTY_JNO & ", " & INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL & ", " & INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                End If
                '==================================================================================================================================================================================
                
                Grid2.Cell(GridImport, 1).Text = 1
            End If
        End If
        I = I + 1
        progCPB.Value = (I / (Grid2.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next
    Prg1.Value = 2
    Call ImportSMISSales
    cmdCheck.Enabled = True: cmdExit.Enabled = True
    Screen.MousePointer = 0
    MsgBox "Import Successfully Completed!", vbInformation, "Finish"
    '=========================================================================================================
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Sub ImportPMISSales()
    Dim PMIOS_TRANTYPE                                 As String
    Dim PMIOS_TRANNO                                   As String
    Dim PMIOS_TRANDATE                                 As String
    Dim PMIOS_cuscde                                   As String
    Dim PMIOS_AcctName                                 As String
    Dim PMIOS_TTLINVAMT                                As Double
    Dim PMIOS_DS_AMT1                                  As Double
    Dim PMIOS_NETINVAMT                                As Double
    Dim PMIOS_NETCOST                                As Double

    Dim PMIOS_TYPE                                     As String
    
    I = 0
    
    For GridImport = 1 To Grid1.Rows - 1
        If N2Str2Zero(Grid1.Cell(GridImport, 1).Text) = 0 Then
            Set rsPMIOS_ORD_HD = New ADODB.Recordset
            If Grid1.Cell(GridImport, 2).Text = "Accessories" Then
               Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where Type = 'A' AND TranType = '" & Left(Grid1.Cell(GridImport, 3).Text, 3) & "' and Tranno = '" & Right(Grid1.Cell(GridImport, 3).Text, 6) & "' AND STATUS = 'P' AND trandate = '" & CDate(dtpTranDate) & "' order by Tranno ASC")
            ElseIf Grid1.Cell(GridImport, 2).Text = "Materials" Then
               Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where Type = 'M' AND TranType = '" & Left(Grid1.Cell(GridImport, 3).Text, 3) & "' and Tranno = '" & Right(Grid1.Cell(GridImport, 3).Text, 6) & "' AND STATUS = 'P' AND trandate = '" & CDate(dtpTranDate) & "' order by Tranno ASC")
            ElseIf Grid1.Cell(GridImport, 2).Text = "Parts" Then
               Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where Type = 'P' AND TranType = '" & Left(Grid1.Cell(GridImport, 3).Text, 3) & "' and Tranno = '" & Right(Grid1.Cell(GridImport, 3).Text, 6) & "' AND STATUS = 'P' AND trandate = '" & CDate(dtpTranDate) & "' order by Tranno ASC")
            Else
               GoTo NextGrid
            End If
            
            If Not rsPMIOS_ORD_HD.EOF And Not rsPMIOS_ORD_HD.BOF Then
                PMIOS_TRANTYPE = Null2String(rsPMIOS_ORD_HD!trantype)
                PMIOS_TRANNO = Null2String(rsPMIOS_ORD_HD!TRANNO)
                PMIOS_TRANDATE = Null2String(rsPMIOS_ORD_HD!trandate)
                PMIOS_cuscde = Null2String(rsPMIOS_ORD_HD!CUSTCODE)
                PMIOS_AcctName = SetCustomerName(rsPMIOS_ORD_HD!CUSTCODE)
                PMIOS_TTLINVAMT = Round(N2Str2Zero(rsPMIOS_ORD_HD!TTLINVAMT), 2)
                PMIOS_DS_AMT1 = Round(N2Str2Zero(rsPMIOS_ORD_HD!DS_AMT1), 2)
                PMIOS_NETINVAMT = Round(N2Str2Zero(rsPMIOS_ORD_HD!NetInvAmt), 2)
                PMIOS_NETCOST = Round(N2Str2Zero(rsPMIOS_ORD_HD!NETCOST), 2)
                PMIOS_TYPE = Null2String(rsPMIOS_ORD_HD!Type)
    
                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If
    
                J_JDATE = N2Date2Null(PMIOS_TRANDATE)
                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                J_JTYPE = "'SJ'"
                J_REMARKS = "NULL"
                J_VENDORCODE = "'999999'"
                J_CUSTOMERCODE = N2Str2Null(PMIOS_cuscde)
    
                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0
    
                J_AMOUNTTOPAY = 0
                J_INVOICEAMT = Round(PMIOS_NETINVAMT, 2)
                J_BALANCE = Round(PMIOS_NETINVAMT, 2)
                J_AMOUNTPAID = 0
    
                J_STATUS = "'N'"
    
                J_INVOICEDATE = N2Date2Null(PMIOS_TRANDATE)
                J_INVOICENO = N2Str2Null(PMIOS_TRANNO)
                J_CHECKNO = "NULL"
                J_DUEDATE = N2Date2Null(PMIOS_TRANDATE)
                J_PAYTYPE = N2Str2Null(rsPMIOS_ORD_HD!TERMS)
                If PMIOS_TYPE = "P" Then
                   J_INVOICETYPE = "'PI'"
                End If
                If PMIOS_TYPE = "A" Then
                   J_INVOICETYPE = "'AI'"
                End If
                If PMIOS_TYPE = "M" Then
                   J_INVOICETYPE = "'MI'"
                End If
                J_CHECKDATE = "NULL"
                J_BANKCODE = "NULL"
                J_REFNO = "NULL"
                J_REFDATE = N2Date2Null(PMIOS_TRANDATE)
                J_TERMS = N2Str2Null(PMIOS_TRANTYPE)
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"
    
                J_JITEMNO = "'0001'"
                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                J_DEBIT = Round(NumericVal(PMIOS_NETINVAMT), 2)
                J_CREDIT = 0
                J_TAX = 0
                J_GROSS = 0
                J_NET = 0
                J_STATUS = "'N'"
                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                J_JITEMNO = "'0002'"
                If PMIOS_TYPE = "P" Then
                   J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS"))
                   J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS")))
                End If
                If PMIOS_TYPE = "A" Then
                   J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("ACCESSORIES"))
                   J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("ACCESSORIES")))
                End If
                If PMIOS_TYPE = "M" Then
                   J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("MATERIALS"))
                   J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("MATERIALS")))
                End If
                J_GROSS = Round(NumericVal(PMIOS_TTLINVAMT), 2)
                J_TAX = Round(NumericVal(Round((PMIOS_TTLINVAMT / 1.12), 2) * 0.12), 2)
                J_NET = Round(NumericVal(PMIOS_TTLINVAMT) - NumericVal(J_TAX), 2)
                J_DEBIT = 0
                J_CREDIT = Round(NumericVal(J_NET), 2)
                J_STATUS = "'N'"
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
    
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                If PMIOS_DS_AMT1 > 0 Then
                    J_JITEMNO = "'0003'"
                    If PMIOS_TYPE = "P" Then
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("PARTS")))
                    End If
                    If PMIOS_TYPE = "A" Then
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("ACCESSORIES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("ACCESSORIES")))
                    End If
                    If PMIOS_TYPE = "M" Then
                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("MATERIALS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("MATERIALS")))
                    End If
                    J_GROSS = Round(NumericVal(PMIOS_DS_AMT1), 2)
                    J_TAX = Round(NumericVal(Round((PMIOS_DS_AMT1 / 1.12), 2) * 0.12), 2)
                    J_NET = Round(NumericVal(PMIOS_DS_AMT1) - NumericVal(J_TAX), 2)
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    J_JITEMNO = "'0004'"
                Else
                    J_JITEMNO = "'0003'"
                End If
                J_ACCT_CODE = N2Str2Null(ReturnOutPutTax())
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutPutTax()))
                J_DEBIT = 0
                J_CREDIT = Round(NumericVal(Round((PMIOS_NETINVAMT / 1.12), 2) * 0.12), 2)
                J_TAX = 0
                J_GROSS = 0
                J_NET = 0
                J_STATUS = "'N'"
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                If Round(Round(TOTAL_CREDIT, 2) - Round(TOTAL_DEBIT, 2), 2) = 0.01 Then J_CREDIT = Round(J_CREDIT, 2) - 0.01
                If Round(Round(TOTAL_DEBIT, 2) - Round(TOTAL_CREDIT, 2), 2) = 0.01 Then J_CREDIT = Round(J_CREDIT, 2) + 0.01
    
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
    
                If J_JITEMNO = "'0004'" Then
                   J_JITEMNO = "'0005'"
                Else
                   J_JITEMNO = "'0004'"
                End If
    
                If PMIOS_TYPE = "P" Then
                   J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("PARTS"))
                   J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("PARTS")))
                End If
                If PMIOS_TYPE = "A" Then
                   J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("ACCESSORIES"))
                   J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("ACCESSORIES")))
                End If
                If PMIOS_TYPE = "M" Then
                   J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("PARTS", "LUBRICANTS"))
                   J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("PARTS", "LUBRICANTS")))
                End If
                J_GROSS = 0
                J_TAX = 0
                J_NET = 0
                J_DEBIT = Round(NumericVal(PMIOS_NETCOST), 2)
                J_CREDIT = 0
                J_STATUS = "'N'"
                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                
                If J_JITEMNO = "'0004'" Then
                   J_JITEMNO = "'0005'"
                Else
                   J_JITEMNO = "'0006'"
                End If
                
                If PMIOS_TYPE = "P" Then
                    J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                End If
                If PMIOS_TYPE = "A" Then
                    J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES", "INVA"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES", "INVA")))
                End If
                If PMIOS_TYPE = "M" Then
                    J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVM"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVM")))
                End If
                J_DEBIT = 0
                J_CREDIT = Round(NumericVal(PMIOS_NETCOST), 2)
                J_TAX = 0
                J_GROSS = 0
                J_NET = 0
                J_STATUS = "'N'"
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
    
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
    
                gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                               " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                 ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                Grid1.Cell(GridImport, 1).Text = 1
            End If
        End If
        I = I + 1
        progCPB.Value = (I / (Grid1.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
NextGrid:
    Next
End Sub

Sub ImportSMISSales()
    Dim rsSMIS_PURCHAGREE                              As ADODB.Recordset
    Dim SMIS_VI_NO                                     As String
    Dim SMIS_DATERELEASED                              As String
    Dim SMIS_CODE                                      As String
    Dim SMIS_AcctName                                  As String
    Dim SMIS_NETSALESPRICE                             As Double
    Dim SMIS_OTHERS                                    As Double
    Dim SMIS_FOB                                       As Double
    Dim SMIS_TOTALCOST                                       As Double
    
    I = 0
    For GridImport = 1 To Grid3.Rows - 1
        If N2Str2Zero(Grid3.Cell(GridImport, 1).Text) = 0 Then
            Set rsSMIS_PURCHAGREE = New ADODB.Recordset
            Set rsSMIS_PURCHAGREE = gconDMIS.Execute("Select * from SMIS_PurchAgree Where VI_NO = '" & Grid3.Cell(GridImport, 3).Text & "' AND CONVERT(VarChar, DateReleased, 101)= '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' order by VI_NO ASC")
            If Not rsSMIS_PURCHAGREE.EOF And Not rsSMIS_PURCHAGREE.BOF Then
                SMIS_VI_NO = Null2String(rsSMIS_PURCHAGREE!VI_NO)
                SMIS_DATERELEASED = Null2String(rsSMIS_PURCHAGREE!DATERELEASED)
                SMIS_CODE = Null2String(rsSMIS_PURCHAGREE!code)
                SMIS_FOB = N2Str2Zero(rsSMIS_PURCHAGREE!FREIGHT)
                SMIS_OTHERS = N2Str2Zero(rsSMIS_PURCHAGREE!OTHERS)
                SMIS_TOTALCOST = N2Str2Zero(rsSMIS_PURCHAGREE!TOTAL_COST)
                If Null2String(rsSMIS_PURCHAGREE!TERM) = "F" Then
                    SMIS_NETSALESPRICE = (N2Str2Zero(rsSMIS_PURCHAGREE!NETSALESPRICE) + SMIS_FOB) - SMIS_OTHERS
                Else
                    SMIS_NETSALESPRICE = N2Str2Zero(rsSMIS_PURCHAGREE!NETSALESPRICE)
                End If
                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If
    
                J_JDATE = N2Date2Null(SMIS_DATERELEASED)
                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                J_JTYPE = "'SJ'"
                J_REMARKS = "NULL"
                J_VENDORCODE = "'999999'"
                J_CUSTOMERCODE = N2Str2Null(SMIS_CODE)
    
                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0
    
                J_AMOUNTTOPAY = 0
                J_INVOICEAMT = Round(SMIS_NETSALESPRICE, 2)
                J_BALANCE = Round(SMIS_NETSALESPRICE, 2)
                J_AMOUNTPAID = 0
    
                J_STATUS = "'N'"
    
                J_INVOICEDATE = N2Date2Null(SMIS_DATERELEASED)
                J_INVOICENO = N2Str2Null(SMIS_VI_NO)
                J_CHECKNO = "NULL"
                J_DUEDATE = N2Date2Null(SMIS_DATERELEASED)
                J_PAYTYPE = "NULL"
                J_INVOICETYPE = "'VI'"
                J_CHECKDATE = "NULL"
                J_BANKCODE = "NULL"
                J_REFNO = "NULL"
                J_REFDATE = N2Date2Null(SMIS_DATERELEASED)
                J_TERMS = "NULL"
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"
    
                J_JITEMNO = "'0001'"
                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SALES"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SALES")))
                If Null2Bool(rsSMIS_PURCHAGREE!ZERORATED) = False Then
                   J_DEBIT = Round(NumericVal(SMIS_NETSALESPRICE), 2)
                Else
                   J_DEBIT = Round(NumericVal(SMIS_NETSALESPRICE), 2)
                End If
                J_CREDIT = 0
                J_TAX = 0
                J_GROSS = 0
                J_NET = 0
                J_STATUS = "'N'"
                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                J_JITEMNO = "'0002'"
                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL)))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL))))
                If Null2Bool(rsSMIS_PURCHAGREE!ZERORATED) = False Then
                   J_GROSS = NumericVal(SMIS_NETSALESPRICE) + NumericVal(SMIS_OTHERS)
                   J_TAX = Round(NumericVal(J_GROSS) / 9.3333, 2)
                   J_NET = Round(NumericVal(J_GROSS) - NumericVal(J_TAX), 2)
                Else
                   J_GROSS = Round(NumericVal(SMIS_NETSALESPRICE) + NumericVal(SMIS_OTHERS), 2)
                   J_TAX = 0
                   J_NET = Round(NumericVal(SMIS_NETSALESPRICE) + NumericVal(SMIS_OTHERS), 2)
                End If
                J_DEBIT = 0
                J_CREDIT = Round(NumericVal(J_NET), 2)
                J_STATUS = "'N'"
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
    
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
    
                J_JITEMNO = "'0003'"
                If SMIS_OTHERS > 0 Then
                    J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL)))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL))))
                    J_GROSS = Round(NumericVal(SMIS_OTHERS), 2)
                    J_TAX = Round(NumericVal(J_GROSS) / 9.3333, 2)
                    J_NET = Round(NumericVal(J_GROSS) - NumericVal(J_TAX), 2)
                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    J_JITEMNO = "'0004'"
                End If
                If Null2Bool(rsSMIS_PURCHAGREE!ZERORATED) = False Then
                    J_ACCT_CODE = N2Str2Null(ReturnOutPutTax())
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutPutTax()))
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(SMIS_NETSALESPRICE / 9.3333), 2)
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    If J_JITEMNO = "'0004'" Then
                        J_JITEMNO = "'0005'"
                    Else
                        J_JITEMNO = "'0004'"
                    End If
                Else
                    If J_JITEMNO = "'0003'" Then
                        J_JITEMNO = "'0004'"
                    Else
                        J_JITEMNO = "'0003'"
                    End If
                End If
                
                J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL)))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL))))
                J_GROSS = 0
                J_TAX = 0
                J_NET = 0
                J_DEBIT = Round(SMIS_TOTALCOST - (NumericVal(SMIS_TOTALCOST) / 9.3333), 2)
                J_CREDIT = 0
                J_STATUS = "'N'"
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
    
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                
                If J_JITEMNO = "'0005'" Then
                    J_JITEMNO = "'0006'"
                Else
                    J_JITEMNO = "'0005'"
                End If
    
                J_ACCT_CODE = N2Str2Null(ReturnInventory("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL)))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL))))
                J_GROSS = 0
                J_TAX = 0
                J_NET = 0
                J_DEBIT = 0
                J_CREDIT = Round(SMIS_TOTALCOST - (NumericVal(SMIS_TOTALCOST) / 9.3333), 2)
                J_STATUS = "'N'"
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
    
                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
    
                gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                               " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                 ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                Grid3.Cell(GridImport, 1).Text = 1
            End If
        End If
        I = I + 1
        progCPB.Value = (I / (Grid3.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next
    Prg1.Value = 3:
End Sub

Sub ProcessPartsCost()
    Dim rsRo_det                                       As ADODB.Recordset
    Dim rsOrd_Hd                                       As ADODB.Recordset
    Dim rsRR_HD As ADODB.Recordset
    Dim rsDAYTRAN                                      As ADODB.Recordset
    Dim rsREPOR                                        As ADODB.Recordset
    Dim rsPartMas                                      As ADODB.Recordset
    Dim vDate_Rel                                      As String
    Dim I                                              As Integer
    Dim IValue As Double
        
    Dim vTDTranDate, vTDTranType, vTDTranno, vSupCode As String
    Dim vTDTranQTY, vTotTranCost, vTotTranInvAmt, vMAC, vPMOnhand, vVatAmt As Double
        
    Set rsTdayTran = New ADODB.Recordset
    rsTdayTran.Open "select id,ItemNo,trantype,tranno,STOCK_ORD,tranqty,status,in_out,tranucost,traninvamt,trandate from PMIS_AllDayTran where tranucost <= 0 AND [TYPE] = 'P' AND trantype <> 'ADB' and (status = 'P' OR status = 'B') and status <> 'N' AND trandate = " & N2Date2Null(dtpTranDate) & " order by id asc", gconDMIS
    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
        rsTdayTran.MoveFirst
        Screen.MousePointer = 11
        I = 0
        Do While Not rsTdayTran.EOF
            gconDMIS.Execute "update PMIS_DayTran set TYPE = 'P', ItemNo = '" & Format(Null2String(rsTdayTran!ItemNo), "0000") & "' where ID = " & rsTdayTran!ID
            gconDMIS.Execute "update PMIS_TDayTran set TYPE = 'P', ItemNo = '" & Format(Null2String(rsTdayTran!ItemNo), "0000") & "' where ID = " & rsTdayTran!ID
            vTDTranDate = N2Date2Null(rsTdayTran!trandate)
            vTDTranType = Null2String(rsTdayTran!trantype)
            vTDTranno = Null2String(rsTdayTran!TRANNO)
            vTDTranQTY = N2Str2Zero(rsTdayTran!tranqty)
            If N2Str2Zero(rsTdayTran!TRANUCOST) > 0 Then
                vTotTranCost = N2Str2Zero(rsTdayTran!TRANUCOST) * vTDTranQTY
            Else
                vTotTranCost = 0
            End If
            If N2Str2Zero(rsTdayTran!TRANINVAMT) > 0 Then
                vTotTranInvAmt = N2Str2Zero(rsTdayTran!TRANINVAMT) * vTDTranQTY
            Else
                vTotTranInvAmt = 0
            End If
            Set rsPartMas = New ADODB.Recordset
            rsPartMas.Open "select id,STOCKNO,mac,tissqty,trecqty,poqty,tprqty,prqty,issuances,receipts,onhand,ONREQUEST,REQSERVED,SERVED,ONORDER,ORDERED,tpoqty,S_REQSERVED,S_ONREQUEST,purchases from PMIS_PARTMAS where STOCKNO = " & N2Str2Null(rsTdayTran!stock_ord), gconDMIS
            If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                vMAC = N2Str2Zero(rsPartMas!Mac)
                vPMOnhand = N2Str2IntZero(rsPartMas!ONHAND)
                If Null2String(rsTdayTran!IN_OUT) = "O" And vTDTranQTY <> 0 Then
                    gconDMIS.Execute "update PMIS_DayTran set tranucost = " & vMAC & " where ID = " & rsTdayTran!ID
                    gconDMIS.Execute "update PMIS_TDayTran set tranucost = " & vMAC & " where ID = " & rsTdayTran!ID
                End If

                If Null2String(rsTdayTran!IN_OUT) = "I" And vTDTranQTY <> 0 Then
                    Set rsRR_HD = New ADODB.Recordset
                    rsRR_HD.Open "Select id,recvd_code,ds1,status,classcode,rrno from PMIS_RR_HD where [TYPE] = 'P' AND rrno = '" & Format(vTDTranno, "000000") & "' and status <> 'C'", gconDMIS
                    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                        vSupCode = Null2String(rsRR_HD!RECVD_CODE)
                        vVatAmt = N2Str2Zero(rsRR_HD!DS1)
                        If Null2String(rsRR_HD!CLASSCODE) = "PCG" Or Null2String(rsRR_HD!CLASSCODE) = "PCS" Then
                            gconDMIS.Execute "update PMIS_RR_HD set ds1 = " & vVatAmt & " where id = " & rsRR_HD!ID
                            If CheckIfNonVatSup(Null2String(rsRR_HD!RECVD_CODE)) = False Then
                                If vSupCode <> VPAMCOR And vVatAmt <= 0 Then
                                    vTotTranCost = vTotTranCost / ConvertToBIRDecimalFormat(vVatAmt)
                                End If
                            Else
                                vTotTranCost = vTotTranInvAmt
                                If N2Str2Zero(rsTdayTran!TRANINVAMT) > 0 Then
                                    gconDMIS.Execute ("update PMIS_DayTran Set tranucost = " & rsTdayTran!TRANINVAMT & " Where id = " & rsTdayTran!ID)
                                    gconDMIS.Execute ("update PMIS_TDayTran Set tranucost = " & rsTdayTran!TRANINVAMT & " Where id = " & rsTdayTran!ID)
                                Else
                                    gconDMIS.Execute ("update PMIS_DayTran Set tranucost = 0 Where id = " & rsTdayTran!ID)
                                    gconDMIS.Execute ("update PMIS_TDayTran Set tranucost = 0 Where id = " & rsTdayTran!ID)
                                End If
                                gconDMIS.Execute ("update PMIS_RR_HD Set DS1 = 0, DS_AMT1 = 0 WHERE ID = " & rsRR_HD!ID)
                                gconDMIS.Execute ("update PMIS_REC_HIST Set DS1 = 0, DS_AMT1 = 0 WHERE ID = " & rsRR_HD!ID)
                            End If
                        End If
                    Else
                        If vTDTranType = "ADJ" Or vTDTranType = "BEG" Then
                           vTotTranCost = vMAC * vTDTranQTY
                        End If
                    End If
                    If vPMOnhand <= 0 Then
                        vMAC = vTotTranCost / vTDTranQTY
                    Else
                        vMAC = ((vMAC * vPMOnhand) + vTotTranCost) / (vTDTranQTY + vPMOnhand)
                    End If
                    gconDMIS.Execute "update PMIS_dayTran set tranucost = " & vMAC & ", mac = " & vMAC & ", status = 'P' where id = " & rsTdayTran!ID
                    gconDMIS.Execute "update PMIS_TdayTran set tranucost = " & vMAC & ", mac = " & vMAC & ", status = 'P' where id = " & rsTdayTran!ID
                End If
            End If
            DoEvents
            I = I + 1
            progCPB.Value = (I / rsTdayTran.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            rsTdayTran.MoveNext
        Loop
        DoEvents
        Screen.MousePointer = 0
    End If
    
    Set rsRo_det = New ADODB.Recordset
    Set rsRo_det = gconDMIS.Execute("Select CSMS_Ro_Det.ID,CSMS_Ro_Det.DetAmt,CSMS_Ro_Det.Rep_Or,detcde,CSMS_repor.Dte_rel from CSMS_Ro_Det inner join CSMS_repor on CSMS_ro_det.rep_or = CSMS_repor.rep_or where CSMS_RO_DET.livil = '2' and CSMS_repor.dte_Comp = " & N2Date2Null(dtpTranDate) & " Order by CSMS_ro_det.Rep_Or asc")
    If Not rsRo_det.EOF And Not rsRo_det.BOF Then
        rsRo_det.MoveFirst
        I = 0
        Do While Not rsRo_det.EOF
            Set rsOrd_Hd = New ADODB.Recordset
            Set rsOrd_Hd = gconDMIS.Execute("Select tranno from PMIS_vw_ISS_HISTORY where [TYPE] = 'P' AND trantype = 'RIV' and RONO = '" & Null2String(rsRo_det!REP_OR) & "'")
            If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                rsOrd_Hd.MoveFirst
                Do While Not rsOrd_Hd.EOF
                    Set rsDAYTRAN = New ADODB.Recordset
                    Set rsDAYTRAN = gconDMIS.Execute("Select tranucost from PMIS_vw_IS_DETHIST where [TYPE] = 'P' AND trantype = 'RIV' and tranno = '" & rsOrd_Hd!TRANNO & "' and STOCK_ORD = " & N2Str2Null(rsRo_det!detcde) & " order by trandate desc")
                    If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
                        If N2Str2Zero(rsDAYTRAN!TRANUCOST) > 0 Then
                            gconDMIS.Execute "Update CSMS_Ro_Det Set " & _
                                             "DetCost = " & N2Str2Zero(rsDAYTRAN!TRANUCOST) & _
                                             " Where id = " & rsRo_det!ID
                        End If
                    End If
                    rsOrd_Hd.MoveNext
                Loop
            End If
            I = I + 1
            progCPB.Value = (I / rsRo_det.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents
            rsRo_det.MoveNext
        Loop
    End If
End Sub

Sub ProcessMaterialsCost()
    Dim rsRo_det                                       As ADODB.Recordset
    Dim rsOrd_Hd                                       As ADODB.Recordset
    Dim rsRR_HD As ADODB.Recordset
    Dim rsDAYTRAN                                      As ADODB.Recordset
    Dim rsREPOR                                        As ADODB.Recordset
    Dim rsPartMas                                      As ADODB.Recordset
    Dim vDate_Rel                                      As String
    Dim I                                              As Integer
    Dim IValue As Double
    
    Dim vTDTranDate, vTDTranType, vTDTranno, vSupCode, vSTOCKDESC As String
    Dim vTDTranQTY, vTotTranCost, vTotTranInvAmt, vMAC, vPMOnhand, vVatAmt As Double

    Set rsTdayTran = New ADODB.Recordset
    rsTdayTran.Open "select id,ItemNo,trantype,tranno,STOCK_ORD,tranqty,status,in_out,tranucost,traninvamt,trandate from PMIS_AllDayTran where tranucost <= 0 AND [TYPE] = 'M' AND trantype <> 'ADB' and (status = 'P' OR status = 'B') and status <> 'N' AND trandate = " & N2Date2Null(dtpTranDate) & " order by id asc", gconDMIS
    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
        rsTdayTran.MoveFirst
        Screen.MousePointer = 11
        I = 0
        Do While Not rsTdayTran.EOF
            gconDMIS.Execute "update PMIS_DayTran set TYPE = 'M', ItemNo = '" & Format(Null2String(rsTdayTran!ItemNo), "0000") & "' where ID = " & rsTdayTran!ID
            gconDMIS.Execute "update PMIS_TDayTran set TYPE = 'M', ItemNo = '" & Format(Null2String(rsTdayTran!ItemNo), "0000") & "' where ID = " & rsTdayTran!ID
            vTDTranDate = N2Date2Null(rsTdayTran!trandate)
            vTDTranType = Null2String(rsTdayTran!trantype)
            vTDTranno = Null2String(rsTdayTran!TRANNO)
            vTDTranQTY = N2Str2IntZero(rsTdayTran!tranqty)
            If N2Str2Zero(rsTdayTran!TRANUCOST) > 0 Then
                vTotTranCost = N2Str2Zero(rsTdayTran!TRANUCOST) * vTDTranQTY
            Else
                vTotTranCost = 0
            End If
            If N2Str2Zero(rsTdayTran!TRANINVAMT) > 0 Then
                vTotTranInvAmt = N2Str2Zero(rsTdayTran!TRANINVAMT) * vTDTranQTY
            Else
                vTotTranInvAmt = 0
            End If
            Set rsPartMas = New ADODB.Recordset
            Set rsPartMas = gconDMIS.Execute("select STOCKNO from PMIS_STOCKMAS where TYPE = 'M' AND STOCKNO = " & N2Str2Null(rsTdayTran!stock_ord))
            If rsPartMas.EOF And rsPartMas.BOF Then
                vSTOCKDESC = "'NO DESCRIPTION'"
                'gconDMIS.Execute ("Insert into PMIS_STOCKMAS ([TYPE],STOCKNO,STOCKDESC,date_entered) values ('M'," & N2Str2Null(rsTdayTran!STOCK_ORD) & "," & vSTOCKDESC & "," & N2Str2Null(rsTdayTran!trandate) & ")")
            Else
                gconDMIS.Execute ("Update PMIS_STOCKMAS SET ACTIVE = 'Y', TYPE = 'M' WHERE STOCKNO = " & N2Str2Null(rsTdayTran!stock_ord))
            End If
            Set rsPartMas = New ADODB.Recordset
            rsPartMas.Open "select id,STOCKNO,mac,tissqty,trecqty,poqty,tprqty,prqty,issuances,receipts,onhand,ONREQUEST,REQSERVED,SERVED,ONORDER,ORDERED,tpoqty,S_REQSERVED,S_ONREQUEST,purchases from PMIS_STOCKMAS where TYPE = 'M' AND STOCKNO = " & N2Str2Null(rsTdayTran!stock_ord), gconDMIS
            If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                vMAC = N2Str2Zero(rsPartMas!Mac)
                vPMOnhand = N2Str2IntZero(rsPartMas!ONHAND)
                If Null2String(rsTdayTran!IN_OUT) = "O" And vTDTranQTY <> 0 Then
                    gconDMIS.Execute "update PMIS_DayTran set tranucost = " & vMAC & " where ID = " & rsTdayTran!ID
                    gconDMIS.Execute "update PMIS_TDayTran set tranucost = " & vMAC & " where ID = " & rsTdayTran!ID
                End If

                If Null2String(rsTdayTran!IN_OUT) = "I" And vTDTranQTY <> 0 Then
                    Set rsRR_HD = New ADODB.Recordset
                    rsRR_HD.Open "Select id,recvd_code,ds1,status,classcode,rrno from PMIS_RR_HD where [TYPE] = 'M' AND rrno = '" & Format(vTDTranno, "000000") & "' and status <> 'C'", gconDMIS
                    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                        vSupCode = Null2String(rsRR_HD!RECVD_CODE)
                        vVatAmt = N2Str2Zero(rsRR_HD!DS1)
                        If Null2String(rsRR_HD!CLASSCODE) = "PCG" Or Null2String(rsRR_HD!CLASSCODE) = "PCS" Then
                            gconDMIS.Execute "update PMIS_RR_HD set ds1 = " & vVatAmt & " where id = " & rsRR_HD!ID
                            If CheckIfNonVatSup(Null2String(rsRR_HD!RECVD_CODE)) = False Then
                                If vSupCode <> VPAMCOR And vVatAmt <= 0 Then
                                    vTotTranCost = vTotTranCost / ConvertToBIRDecimalFormat(vVatAmt)
                                End If
                            Else
                                vTotTranCost = vTotTranInvAmt
                                If N2Str2Zero(rsTdayTran!TRANINVAMT) > 0 Then
                                    gconDMIS.Execute ("update PMIS_DayTran Set tranucost = " & rsTdayTran!TRANINVAMT & " Where id = " & rsTdayTran!ID)
                                    gconDMIS.Execute ("update PMIS_TDayTran Set tranucost = " & rsTdayTran!TRANINVAMT & " Where id = " & rsTdayTran!ID)
                                Else
                                    gconDMIS.Execute ("update PMIS_DayTran Set tranucost = 0 Where id = " & rsTdayTran!ID)
                                    gconDMIS.Execute ("update PMIS_TDayTran Set tranucost = 0 Where id = " & rsTdayTran!ID)
                                End If
                                gconDMIS.Execute ("update PMIS_RR_HD Set DS1 = 0, DS_AMT1 = 0 WHERE ID = " & rsRR_HD!ID)
                                gconDMIS.Execute ("update PMIS_REC_HIST Set DS1 = 0, DS_AMT1 = 0 WHERE ID = " & rsRR_HD!ID)
                            End If
                        End If
                    Else
                        If vTDTranType = "ADJ" Or vTDTranType = "BEG" Then
                           vTotTranCost = vMAC * vTDTranQTY
                        End If
                    End If
                    If vPMOnhand <= 0 Then
                        vMAC = vTotTranCost / vTDTranQTY
                    Else
                        vMAC = ((vMAC * vPMOnhand) + vTotTranCost) / (vTDTranQTY + vPMOnhand)
                    End If
                    gconDMIS.Execute "update PMIS_dayTran set tranucost = " & vMAC & ", mac = " & vMAC & ", status = 'P' where id = " & rsTdayTran!ID
                    gconDMIS.Execute "update PMIS_TdayTran set tranucost = " & vMAC & ", mac = " & vMAC & ", status = 'P' where id = " & rsTdayTran!ID
                End If
            End If
            DoEvents
            I = I + 1
            progCPB.Value = (I / rsTdayTran.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsTdayTran.MoveNext
        Loop
        DoEvents
        Screen.MousePointer = 0
    End If
        
    
    Set rsRo_det = New ADODB.Recordset
    Set rsRo_det = gconDMIS.Execute("Select CSMS_Ro_Det.ID,CSMS_Ro_Det.DetAmt,CSMS_Ro_Det.Rep_Or,detcde,CSMS_repor.Dte_rel from CSMS_Ro_Det inner join CSMS_repor on CSMS_ro_det.rep_or = CSMS_repor.rep_or where CSMS_RO_DET.livil = '3' and CSMS_repor.dte_Comp = " & N2Date2Null(dtpTranDate) & " Order by CSMS_ro_det.Rep_Or asc")
    If Not rsRo_det.EOF And Not rsRo_det.BOF Then
        rsRo_det.MoveFirst
        I = 0
        Do While Not rsRo_det.EOF
            Set rsOrd_Hd = New ADODB.Recordset
            Set rsOrd_Hd = gconDMIS.Execute("Select tranno from PMIS_vw_ISS_HISTORY where [TYPE] = 'M' AND trantype = 'RIV' and RONO = '" & Null2String(rsRo_det!REP_OR) & "'")
            If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                rsOrd_Hd.MoveFirst
                Do While Not rsOrd_Hd.EOF
                    Set rsDAYTRAN = New ADODB.Recordset
                    Set rsDAYTRAN = gconDMIS.Execute("Select tranucost from PMIS_vw_IS_DETHIST where [TYPE] = 'M' AND trantype = 'RIV' and tranno = '" & rsOrd_Hd!TRANNO & "' and STOCK_ORD = " & N2Str2Null(rsRo_det!detcde) & " order by trandate desc")
                    If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
                        If N2Str2Zero(rsDAYTRAN!TRANUCOST) > 0 Then
                            gconDMIS.Execute "Update CSMS_Ro_Det Set " & _
                                             "DetCost = " & N2Str2Zero(rsDAYTRAN!TRANUCOST) & _
                                           " Where id = " & rsRo_det!ID
                        End If
                    End If
                    rsOrd_Hd.MoveNext
                Loop
            End If
            I = I + 1
            progCPB.Value = (I / rsRo_det.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents
            rsRo_det.MoveNext
        Loop
    End If
End Sub

Sub ProcessAccessoriesCost()
    Dim rsRo_det                                       As ADODB.Recordset
    Dim rsOrd_Hd                                       As ADODB.Recordset
    Dim rsRR_HD As ADODB.Recordset
    Dim rsDAYTRAN                                      As ADODB.Recordset
    Dim rsREPOR                                        As ADODB.Recordset
    Dim rsPartMas                                      As ADODB.Recordset
    Dim vDate_Rel                                      As String
    Dim I                                              As Integer
    Dim IValue As Double
        
    Dim vTDTranDate, vTDTranType, vTDTranno, vSupCode, vSTOCKDESC As String
    Dim vTDTranQTY, vTotTranCost, vTotTranInvAmt, vMAC, vPMOnhand, vVatAmt As Double
    
    Set rsTdayTran = New ADODB.Recordset
    rsTdayTran.Open "select id,ItemNo,trantype,tranno,STOCK_ORD,tranqty,status,in_out,tranucost,traninvamt,trandate from PMIS_AllDayTran where tranucost <= 0 AND [TYPE] = 'A' AND trantype <> 'ADB' and (status = 'P' OR status = 'B') and status <> 'N' AND trandate = " & N2Date2Null(dtpTranDate) & " order by id asc", gconDMIS
    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
        rsTdayTran.MoveFirst
        Screen.MousePointer = 11
        I = 0
        Do While Not rsTdayTran.EOF
            gconDMIS.Execute "update PMIS_DayTran set TYPE = 'A', ItemNo = '" & Format(Null2String(rsTdayTran!ItemNo), "0000") & "' where ID = " & rsTdayTran!ID
            gconDMIS.Execute "update PMIS_TDayTran set TYPE = 'A', ItemNo = '" & Format(Null2String(rsTdayTran!ItemNo), "0000") & "' where ID = " & rsTdayTran!ID
            vTDTranDate = N2Date2Null(rsTdayTran!trandate)
            vTDTranType = Null2String(rsTdayTran!trantype)
            vTDTranno = Null2String(rsTdayTran!TRANNO)
            vTDTranQTY = N2Str2IntZero(rsTdayTran!tranqty)
            If N2Str2Zero(rsTdayTran!TRANUCOST) > 0 Then
                vTotTranCost = N2Str2Zero(rsTdayTran!TRANUCOST) * vTDTranQTY
            Else
                vTotTranCost = 0
            End If
            If N2Str2Zero(rsTdayTran!TRANINVAMT) > 0 Then
                vTotTranInvAmt = N2Str2Zero(rsTdayTran!TRANINVAMT) * vTDTranQTY
            Else
                vTotTranInvAmt = 0
            End If
            Set rsPartMas = New ADODB.Recordset
            Set rsPartMas = gconDMIS.Execute("select STOCKNO from PMIS_STOCKMAS where TYPE = 'A' AND STOCKNO = " & N2Str2Null(rsTdayTran!stock_ord))
            If rsPartMas.EOF And rsPartMas.BOF Then
                vSTOCKDESC = "'NO DESCRIPTION'"
                'gconDMIS.Execute ("Insert into PMIS_STOCKMAS ([TYPE],STOCKNO,STOCKDESC,date_entered) values ('A'," & N2Str2Null(rsTdayTran!STOCK_ORD) & "," & vSTOCKDESC & "," & N2Str2Null(rsTdayTran!trandate) & ")")
            Else
                gconDMIS.Execute ("Update PMIS_STOCKMAS SET ACTIVE = 'Y', TYPE = 'A' WHERE STOCKNO = " & N2Str2Null(rsTdayTran!stock_ord))
            End If
            Set rsPartMas = New ADODB.Recordset
            rsPartMas.Open "select id,STOCKNO,mac,tissqty,trecqty,poqty,tprqty,prqty,issuances,receipts,onhand,ONREQUEST,REQSERVED,SERVED,ONORDER,ORDERED,tpoqty,S_REQSERVED,S_ONREQUEST,purchases from PMIS_STOCKMAS where TYPE = 'A' AND STOCKNO = " & N2Str2Null(rsTdayTran!stock_ord), gconDMIS
            If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                vMAC = N2Str2Zero(rsPartMas!Mac)
                vPMOnhand = N2Str2IntZero(rsPartMas!ONHAND)
                If Null2String(rsTdayTran!IN_OUT) = "O" And vTDTranQTY <> 0 Then
                    gconDMIS.Execute "update PMIS_DayTran set tranucost = " & vMAC & " where ID = " & rsTdayTran!ID
                    gconDMIS.Execute "update PMIS_TDayTran set tranucost = " & vMAC & " where ID = " & rsTdayTran!ID
                End If

                If Null2String(rsTdayTran!IN_OUT) = "I" And vTDTranQTY <> 0 Then
                    Set rsRR_HD = New ADODB.Recordset
                    rsRR_HD.Open "Select id,recvd_code,ds1,status,classcode,rrno from PMIS_RR_HD where [TYPE] = 'A' AND rrno = '" & Format(vTDTranno, "000000") & "' and status <> 'C'", gconDMIS
                    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                        vSupCode = Null2String(rsRR_HD!RECVD_CODE)
                        vVatAmt = N2Str2Zero(rsRR_HD!DS1)
                        If Null2String(rsRR_HD!CLASSCODE) = "PCG" Or Null2String(rsRR_HD!CLASSCODE) = "PCS" Then
                            gconDMIS.Execute "update PMIS_RR_HD set ds1 = " & vVatAmt & " where id = " & rsRR_HD!ID
                            If CheckIfNonVatSup(Null2String(rsRR_HD!RECVD_CODE)) = False Then
                                If vSupCode <> VPAMCOR And vVatAmt <= 0 Then
                                    vTotTranCost = vTotTranCost / ConvertToBIRDecimalFormat(vVatAmt)
                                End If
                            Else
                                vTotTranCost = vTotTranInvAmt
                                If N2Str2Zero(rsTdayTran!TRANINVAMT) > 0 Then
                                    gconDMIS.Execute ("update PMIS_DayTran Set tranucost = " & rsTdayTran!TRANINVAMT & " Where id = " & rsTdayTran!ID)
                                    gconDMIS.Execute ("update PMIS_TDayTran Set tranucost = " & rsTdayTran!TRANINVAMT & " Where id = " & rsTdayTran!ID)
                                Else
                                    gconDMIS.Execute ("update PMIS_DayTran Set tranucost = 0 Where id = " & rsTdayTran!ID)
                                    gconDMIS.Execute ("update PMIS_TDayTran Set tranucost = 0 Where id = " & rsTdayTran!ID)
                                End If
                                gconDMIS.Execute ("update PMIS_RR_HD Set DS1 = 0, DS_AMT1 = 0 WHERE ID = " & rsRR_HD!ID)
                                gconDMIS.Execute ("update PMIS_REC_HIST Set DS1 = 0, DS_AMT1 = 0 WHERE ID = " & rsRR_HD!ID)
                            End If
                        End If
                    Else
                        If vTDTranType = "ADJ" Or vTDTranType = "BEG" Then
                           vTotTranCost = vMAC * vTDTranQTY
                        End If
                    End If
                    If vPMOnhand <= 0 Then
                        vMAC = vTotTranCost / vTDTranQTY
                    Else
                        vMAC = ((vMAC * vPMOnhand) + vTotTranCost) / (vTDTranQTY + vPMOnhand)
                    End If
                    gconDMIS.Execute "update PMIS_dayTran set tranucost = " & vMAC & ", mac = " & vMAC & ", status = 'P' where id = " & rsTdayTran!ID
                    gconDMIS.Execute "update PMIS_TdayTran set tranucost = " & vMAC & ", mac = " & vMAC & ", status = 'P' where id = " & rsTdayTran!ID
                End If
            End If
            DoEvents
            I = I + 1
            progCPB.Value = (I / rsTdayTran.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsTdayTran.MoveNext
        Loop
        DoEvents
        Screen.MousePointer = 0
    End If
        
        
    Set rsRo_det = New ADODB.Recordset
    Set rsRo_det = gconDMIS.Execute("Select CSMS_Ro_Det.ID,CSMS_Ro_Det.DetAmt,CSMS_Ro_Det.Rep_Or,detcde,CSMS_repor.Dte_rel from CSMS_Ro_Det inner join CSMS_repor on CSMS_ro_det.rep_or = CSMS_repor.rep_or where CSMS_RO_DET.livil = '4' and CSMS_repor.dte_Comp = " & N2Date2Null(dtpTranDate) & " Order by CSMS_ro_det.Rep_Or asc")
    If Not rsRo_det.EOF And Not rsRo_det.BOF Then
        rsRo_det.MoveFirst
        I = 0
        Do While Not rsRo_det.EOF
            Set rsOrd_Hd = New ADODB.Recordset
            Set rsOrd_Hd = gconDMIS.Execute("Select tranno from PMIS_vw_ISS_HISTORY where [TYPE] = 'A' AND trantype = 'RIV' and RONO = '" & Null2String(rsRo_det!REP_OR) & "'")
            If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                rsOrd_Hd.MoveFirst
                Do While Not rsOrd_Hd.EOF
                    Set rsDAYTRAN = New ADODB.Recordset
                    Set rsDAYTRAN = gconDMIS.Execute("Select tranucost from PMIS_vw_IS_DETHIST where [TYPE] = 'A' AND trantype = 'RIV' and tranno = '" & rsOrd_Hd!TRANNO & "' and STOCK_ORD = " & N2Str2Null(rsRo_det!detcde) & " order by trandate desc")
                    If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
                        If N2Str2Zero(rsDAYTRAN!TRANUCOST) > 0 Then
                            gconDMIS.Execute "Update CSMS_Ro_Det Set " & _
                                             "DetCost = " & N2Str2Zero(rsDAYTRAN!TRANUCOST) & _
                                           " Where id = " & rsRo_det!ID
                        End If
                    End If
                    rsOrd_Hd.MoveNext
                Loop
            End If
            I = I + 1
            progCPB.Value = (I / rsRo_det.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents
            rsRo_det.MoveNext
        Loop
    End If
End Sub

Private Sub cmdClearJournals_Click()
    Dim rsCHATCheckControlIfExistRecordInJournalHD     As ADODB.Recordset
    Set rsCHATCheckControlIfExistRecordInJournalHD = New ADODB.Recordset
    Set rsCHATCheckControlIfExistRecordInJournalHD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'SJ' and Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
    If Not rsCHATCheckControlIfExistRecordInJournalHD.EOF And Not rsCHATCheckControlIfExistRecordInJournalHD.BOF Then
        Screen.MousePointer = 0
        If MsgBox("Clear Unposted Data for this Particular Date?", vbQuestion + vbYesNo, "Purge Data") = vbYes Then
           Screen.MousePointer = 11
           gconDMIS.Execute ("delete from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'SJ' and Jdate = '" & CDate(dtpTranDate) & "'")
           gconDMIS.Execute ("delete from AMIS_Journal_DET Where STATUS <> 'P' AND Jtype = 'SJ' and Jdate = '" & CDate(dtpTranDate) & "'")
           cmdShowTrans.Value = True
           Screen.MousePointer = 0
           MsgBox "Existing Data Successfully deleted.", vbInformation, "Deleted"
           Exit Sub
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Function SetCustomerName(VVV As Variant) As String
    Dim rsCustomer2                                    As ADODB.Recordset
    Set rsCustomer2 = New ADODB.Recordset
    rsCustomer2.Open "Select CustCode,acctname from ALL_CUSTMASTER_AMIS where CustCode = '" & VVV & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCustomer2.EOF And Not rsCustomer2.BOF Then
        SetCustomerName = UCase(Null2String(rsCustomer2!ACCTname))
    Else
        SetCustomerName = ""
    End If
End Function

Function Setacctname(VVV As String) As String
    Dim rsChartAccount2                                As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    If Left(VVV, 1) = "'" Then
       rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = " & VVV, gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
       rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = '" & VVV & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        Setacctname = UCase(Null2String(rsChartAccount2!Description))
    Else
        Setacctname = ""
    End If
End Function

Function GetVoucherNo() As String
    Dim rsJournal_HD                                   As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'SJ' Order by MAX_VOUCHERNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetVoucherNo = Format(NumericVal(rsJournal_HD!MAX_VOUCHERNO) + 1, "000000")
    Else
        GetVoucherNo = "000001"
    End If
End Function

Private Sub cmdShowTrans_Click()
Screen.MousePointer = 11
InitGrids: DoEvents: cmdCheck.Enabled = False: cmdClearJournals.Enabled = False
Grid1.Rows = 2: Grid2.Rows = 2: Grid3.Rows = 2
Dim InvoiceType, InvoiceTypeCode As String
Dim IS_Exist As Byte
Dim KIM As Integer
Set rsPMIOS_ORD_HD = New ADODB.Recordset
Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where (TranType = 'CSH' OR  TranType = 'CHG') and STATUS = 'P' AND trandate = '" & CDate(dtpTranDate) & "' order by Tranno ASC")
If Not rsPMIOS_ORD_HD.EOF And Not rsPMIOS_ORD_HD.BOF Then
   rsPMIOS_ORD_HD.MoveFirst: KIM = 0
   Grid1.AutoRedraw = False
   Do While Not rsPMIOS_ORD_HD.EOF
      KIM = KIM + 1
      If Null2String(rsPMIOS_ORD_HD!Type) = "P" Then
         InvoiceType = "Parts"
         InvoiceTypeCode = "PI"
      ElseIf Null2String(rsPMIOS_ORD_HD!Type) = "A" Then
         InvoiceType = "Accessories"
         InvoiceTypeCode = "AI"
      ElseIf Null2String(rsPMIOS_ORD_HD!Type) = "M" Then
         InvoiceType = "Materials"
         InvoiceTypeCode = "MI"
      Else
         InvoiceType = "Unknown"
         InvoiceTypeCode = ""
      End If
      If CheckPMISSJExisting(InvoiceTypeCode, Null2String(rsPMIOS_ORD_HD!TRANNO), (rsPMIOS_ORD_HD!trantype)) = True Then
         IS_Exist = 1
      Else
         IS_Exist = 0
      End If
      Grid1.AddItem IS_Exist & Chr(9) & InvoiceType & Chr(9) & Null2String(rsPMIOS_ORD_HD!trantype) & "-" & Null2String(rsPMIOS_ORD_HD!TRANNO) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsPMIOS_ORD_HD!NetInvAmt)) & Chr(9) & Null2String(rsPMIOS_ORD_HD!CUSTNAME)
      
      rsPMIOS_ORD_HD.MoveNext
   Loop
   If KIM > 0 Then Grid1.RemoveItem 1
   Grid1.AutoRedraw = True
   Grid1.Refresh
End If
Set rsCSMIOS_REPOR = New ADODB.Recordset
Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where invoice <> 'NO CHG' AND invoice <> 'PDI RO' and dte_comp = '" & CDate(dtpTranDate) & "' order by invoice ASC")
If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
   rsCSMIOS_REPOR.MoveFirst: KIM = 0
   Grid2.AutoRedraw = False
   Do While Not rsCSMIOS_REPOR.EOF
      KIM = KIM + 1
      If CheckSJExisting("SI", Null2String(rsCSMIOS_REPOR!Invoice)) = True Then
         IS_Exist = 1
      Else
         IS_Exist = 0
      End If
      Grid2.AddItem IS_Exist & Chr(9) & Null2String(rsCSMIOS_REPOR!REP_OR) & Chr(9) & Null2String(rsCSMIOS_REPOR!Invoice) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCSMIOS_REPOR!amount)) & Chr(9) & Null2String(rsCSMIOS_REPOR!Niym)
      rsCSMIOS_REPOR.MoveNext
   Loop
   If KIM > 0 Then Grid2.RemoveItem 1
   Grid2.AutoRedraw = True
   Grid2.Refresh
End If
Set rsSMIS_PURCHAGREE = New ADODB.Recordset

Set rsSMIS_PURCHAGREE = gconDMIS.Execute("Select * from SMIS_PurchAgree Where STATUS = 'P' AND CONVERT(VarChar, DateReleased, 101)  = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' order by VI_NO ASC")
'Set rsSMIS_PURCHAGREE = gconDMIS.Execute("Select * from SMIS_PurchAgree Where dateRELEASED = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' order by VI_NO ASC")
If Not rsSMIS_PURCHAGREE.EOF And Not rsSMIS_PURCHAGREE.BOF Then
   rsSMIS_PURCHAGREE.MoveFirst: KIM = 0
   Grid3.AutoRedraw = False
   Do While Not rsSMIS_PURCHAGREE.EOF
      KIM = KIM + 1
      If CheckSJExisting("VI", Null2String(rsSMIS_PURCHAGREE!VI_NO)) = True Then
         IS_Exist = 1
      Else
         IS_Exist = 0
      End If
      Grid3.AddItem IS_Exist & Chr(9) & Null2String(rsSMIS_PURCHAGREE!IGNKEY_NO) & Chr(9) & Null2String(rsSMIS_PURCHAGREE!VI_NO) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsSMIS_PURCHAGREE!TOTAL)) & Chr(9) & SetCustomerName(Null2String(rsSMIS_PURCHAGREE!code))
      rsSMIS_PURCHAGREE.MoveNext
   Loop
   If KIM > 0 Then Grid3.RemoveItem 1
   Grid3.AutoRedraw = True
   Grid3.Refresh
End If
If KIM > 0 Then
   cmdCheck.Enabled = True
   cmdClearJournals.Enabled = True
End If
Screen.MousePointer = 0
End Sub
Function CheckSJExisting(VarInvoiceType As String, VarInvoiceNo As String) As Boolean
    Dim rsCheckSJ_Journal_HD                                          As ADODB.Recordset
    Set rsCheckSJ_Journal_HD = New ADODB.Recordset
    Set rsCheckSJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'SJ' AND Status <> 'C' AND InvoiceType = " & N2Str2Null(VarInvoiceType) & " AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    If Not rsCheckSJ_Journal_HD.EOF And Not rsCheckSJ_Journal_HD.BOF Then
        CheckSJExisting = True
    Else
        CheckSJExisting = False
    End If
    Set rsCheckSJ_Journal_HD = Nothing
End Function
Private Sub dtpTranDate_Change()
InitGrids: DoEvents:
Grid1.Rows = 1
Grid2.Rows = 1
Grid3.Rows = 1
cmdCheck.Enabled = False
cmdClearJournals.Enabled = False
End Sub

Private Sub Form_Load()
    On Error GoTo Errorcode
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpTranDate = LOGDATE
    InitGrids
    cmdCheck.Enabled = False
    Screen.MousePointer = 0
    Exit Sub

Errorcode:
    Screen.MousePointer = 0
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Database Connection Error!"
    Unload frmSplash
    cmdCheck.Enabled = False
End Sub

Sub InitGrids()
With Grid1
    .Rows = 1
    .Cell(0, 1).Text = "Imported"
    .Cell(0, 2).Text = "Inv. Type"
    .Cell(0, 3).Text = "Inv. No."
    .Cell(0, 4).Text = "Inv. Amt."
    .Cell(0, 5).Text = "Customer"
    
    .Column(0).Width = 10
    .Column(1).Width = 50
    .Column(2).Width = 80
    .Column(3).Width = 80
    .Column(4).Width = 80
    .Column(5).Width = 200

    .Column(1).CellType = cellCheckBox
    .Column(4).Alignment = cellRightGeneral

    .Column(1).Locked = True
    .Column(2).Locked = True
    .Column(3).Locked = True
    .Column(4).Locked = True
    .Column(5).Locked = True

End With

With Grid2
    .Rows = 1
    .Cell(0, 1).Text = "Imported"
    .Cell(0, 2).Text = "RO No."
    .Cell(0, 3).Text = "Inv. No."
    .Cell(0, 4).Text = "Inv. Amt."
    .Cell(0, 5).Text = "Customer"
    
    .Column(0).Width = 10
    .Column(1).Width = 50
    .Column(2).Width = 80
    .Column(3).Width = 60
    .Column(4).Width = 80
    .Column(5).Width = 200

    .Column(1).CellType = cellCheckBox
    .Column(4).Alignment = cellRightGeneral
    
    .Column(1).Locked = True
    .Column(2).Locked = True
    .Column(3).Locked = True
    .Column(4).Locked = True
    .Column(5).Locked = True
End With

With Grid3
    .Rows = 1
    .Cell(0, 1).Text = "Imported"
    .Cell(0, 2).Text = "CS No."
    .Cell(0, 3).Text = "Inv. No."
    .Cell(0, 4).Text = "Inv. Amt."
    .Cell(0, 5).Text = "Customer"
    
    .Column(0).Width = 10
    .Column(1).Width = 50
    .Column(2).Width = 80
    .Column(3).Width = 80
    .Column(4).Width = 80
    .Column(5).Width = 200

    .Column(1).CellType = cellCheckBox
    .Column(4).Alignment = cellRightGeneral

    .Column(1).Locked = True
    .Column(2).Locked = True
    .Column(3).Locked = True
    .Column(4).Locked = True
    .Column(5).Locked = True
End With
DoEvents
End Sub

Function ReturnAR_AccountCode(XXX As String) As String
Dim rsChartAccount As ADODB.Recordset
Set rsChartAccount = New ADODB.Recordset
Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'AR' AND TRANTYPE1 = '" & XXX & "'")
If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
    ReturnAR_AccountCode = Null2String(rsChartAccount!AcctCode)
End If
Set rsChartAccount = Nothing
End Function

Function ReturnSales_AccountCode(INVTYPE As String, Optional OTHERTYPE As String) As String
Dim rsChartAccount As ADODB.Recordset
Set rsChartAccount = New ADODB.Recordset
If OTHERTYPE = "" Then
   Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'SALES' AND TRANTYPE2 = '" & INVTYPE & "'")
Else
   Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'SALES' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE1 = '" & OTHERTYPE & "'")
End If
If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
    ReturnSales_AccountCode = Null2String(rsChartAccount!AcctCode)
End If
Set rsChartAccount = Nothing
End Function

Function ReturnDiscount_AccountCode(INVTYPE As String, Optional OTHERTYPE As String) As String
Dim rsChartAccount As ADODB.Recordset
Set rsChartAccount = New ADODB.Recordset
If OTHERTYPE = "" Then
   Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'DISCOUNT' AND TRANTYPE2 = '" & INVTYPE & "'")
Else
   Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'DISCOUNT' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE1 = '" & OTHERTYPE & "'")
End If
If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
    ReturnDiscount_AccountCode = Null2String(rsChartAccount!AcctCode)
End If
Set rsChartAccount = Nothing
End Function

Function ReturnCostOfSales(INVTYPE As String, Optional OTHERTYPE As String) As String
Dim rsChartAccount As ADODB.Recordset
Set rsChartAccount = New ADODB.Recordset
If OTHERTYPE = "" Then
   Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'COST OF SALES' AND TRANTYPE2 = '" & INVTYPE & "'")
Else
   Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'COST OF SALES' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE1 = '" & OTHERTYPE & "'")
End If
If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
    ReturnCostOfSales = Null2String(rsChartAccount!AcctCode)
End If
Set rsChartAccount = Nothing
End Function

Function ReturnInventory(INVTYPE As String, Optional OTHERTYPE As String) As String
Dim rsChartAccount As ADODB.Recordset
Set rsChartAccount = New ADODB.Recordset
If OTHERTYPE = "" Then
   Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'INVENTORY' AND TRANTYPE2 = '" & INVTYPE & "'")
Else
   Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'INVENTORY' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE1 = '" & OTHERTYPE & "'")
End If
If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
    ReturnInventory = Null2String(rsChartAccount!AcctCode)
End If
Set rsChartAccount = Nothing
End Function

Function ReturnOutPutTax()
Dim rsChartAccount As ADODB.Recordset
Set rsChartAccount = New ADODB.Recordset
Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'OUTPUT TAX'")
If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
    ReturnOutPutTax = Null2String(rsChartAccount!AcctCode)
End If
Set rsChartAccount = Nothing
End Function

Function SetVendorName(VVV As Variant)
    Dim rsVENDOR As ADODB.Recordset
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "Select code,nameofvendor from ALL_Vendor where code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorName = Null2String(rsVENDOR!nameofvendor)
    Else
        SetVendorName = ""
    End If
    Set rsVENDOR = New ADODB.Recordset
End Function

Function CheckIfNonVatSup(SupplierCode As String) As Boolean
    Dim rsSupplierMaster                               As ADODB.Recordset
    Set rsSupplierMaster = New ADODB.Recordset
    rsSupplierMaster.Open "Select supcode,supname,NONVAT from PMIS_vw_Supplier where supcode = '" & SupplierCode & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSupplierMaster.EOF And Not rsSupplierMaster.BOF Then
        If Null2String(rsSupplierMaster!NONVAT) = "Y" Then
            CheckIfNonVatSup = True
        Else
            CheckIfNonVatSup = False
        End If
    Else
        CheckIfNonVatSup = False
    End If
End Function
Function CheckPMISSJExisting(VarInvoiceType As String, VarInvoiceNo As String, VarTerms) As Boolean
    'Update By BTT : 09-24-20008
    Dim rsCheckSJ_Journal_HD                                          As ADODB.Recordset
    Set rsCheckSJ_Journal_HD = New ADODB.Recordset
    Set rsCheckSJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'SJ' and terms = '" & VarTerms & "' AND Status <> 'C' AND InvoiceType = " & N2Str2Null(VarInvoiceType) & " AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    If Not rsCheckSJ_Journal_HD.EOF And Not rsCheckSJ_Journal_HD.BOF Then
        CheckPMISSJExisting = True
    Else
        CheckPMISSJExisting = False
    End If
    Set rsCheckSJ_Journal_HD = Nothing
End Function


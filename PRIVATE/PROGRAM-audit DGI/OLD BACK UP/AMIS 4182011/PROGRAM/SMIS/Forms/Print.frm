VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmSMIS_Report_Print 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing Process..."
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4290
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FCFCFC&
   Icon            =   "Print.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   4290
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   5745
      Left            =   -30
      ScaleHeight     =   5745
      ScaleWidth      =   13290
      TabIndex        =   0
      Top             =   30
      Width           =   13290
      Begin wizButton.cmd cmdDebitMemo 
         Height          =   435
         Left            =   150
         TabIndex        =   5
         ToolTipText     =   "PNP Motor Vehicle Clearance Application"
         Top             =   2149
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "Debit Memo"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print.frx":0E42
      End
      Begin Crystal.CrystalReport rptPrint 
         Left            =   390
         Top             =   570
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin wizButton.cmd cmdCreditMemo 
         Height          =   435
         Left            =   150
         TabIndex        =   4
         ToolTipText     =   "Release Order"
         Top             =   1653
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "Credit Memo"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print.frx":0E5E
      End
      Begin wizButton.cmd cmdVDR 
         Height          =   435
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Vehicle Delivery Report"
         Top             =   660
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "&Vehicle Delivery Report"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print.frx":0E7A
      End
      Begin wizButton.cmd cmdVI 
         Height          =   435
         Left            =   150
         TabIndex        =   1
         ToolTipText     =   "Vehicle Invoice"
         Top             =   165
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "Vehicle &Invoice"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print.frx":0E96
      End
      Begin wizButton.cmd cmdExit 
         Height          =   435
         Left            =   150
         TabIndex        =   6
         ToolTipText     =   "Exit"
         Top             =   5100
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "E&xit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print.frx":0EB2
      End
      Begin wizButton.cmd cmd1 
         Height          =   435
         Left            =   150
         TabIndex        =   3
         ToolTipText     =   "Gate Pass"
         Top             =   1157
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "&Gate Pass"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print.frx":0ECE
      End
      Begin wizButton.cmd cmdReleaseOrder 
         Height          =   435
         Left            =   150
         TabIndex        =   7
         ToolTipText     =   "PNP Motor Vehicle Clearance Application"
         Top             =   2645
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "Release Order"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print.frx":0EEA
      End
      Begin wizButton.cmd cmdClearance 
         Height          =   435
         Left            =   4500
         TabIndex        =   8
         ToolTipText     =   "PNP Motor Vehicle Clearance Application"
         Top             =   4560
         Visible         =   0   'False
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "Clearance Report"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print.frx":0F06
      End
      Begin wizButton.cmd cmdDR 
         Height          =   435
         Left            =   4500
         TabIndex        =   9
         Top             =   5040
         Visible         =   0   'False
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "&Dealers Report"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print.frx":0F22
      End
      Begin wizButton.cmd cmd4 
         Height          =   435
         Left            =   150
         TabIndex        =   10
         ToolTipText     =   "PNP Motor Vehicle Clearance Application"
         Top             =   3141
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "Warranty"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print.frx":0F3E
      End
      Begin wizButton.cmd cmd6 
         Height          =   435
         Left            =   150
         TabIndex        =   11
         ToolTipText     =   "PNP Motor Vehicle Clearance Application"
         Top             =   3637
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "Signatories"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print.frx":0F5A
      End
      Begin wizButton.cmd cmdTransaction 
         Height          =   435
         Left            =   150
         TabIndex        =   12
         ToolTipText     =   "PNP Motor Vehicle Clearance Application"
         Top             =   4140
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "Transaction Slip"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print.frx":0F76
      End
      Begin wizButton.cmd cmdjob 
         Height          =   435
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "PNP Motor Vehicle Clearance Application"
         Top             =   4620
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   767
         TX              =   "Job Request Form"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print.frx":0F92
      End
   End
   Begin VB.PictureBox picDebitMemo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   570
      ScaleHeight     =   2085
      ScaleWidth      =   2865
      TabIndex        =   20
      Top             =   1590
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtDebitMemo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         MaxLength       =   6
         TabIndex        =   21
         ToolTipText     =   "Input Debit Memo Serial Number"
         Top             =   690
         Width           =   2565
      End
      Begin wizButton.cmd cmdPrintDebitmemo 
         Height          =   435
         Left            =   480
         TabIndex        =   22
         Top             =   1320
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         TX              =   "&Print"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print.frx":0FAE
      End
      Begin wizButton.cmd cmd7 
         Height          =   435
         Left            =   1590
         TabIndex        =   23
         Top             =   1320
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         TX              =   "&Back"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print.frx":0FCA
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   315
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   2865
         _Version        =   655364
         _ExtentX        =   5054
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Debit Memo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         Alignment       =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DEBIT MEMO NO."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   450
         Width           =   1695
      End
   End
   Begin VB.PictureBox picCreditMemo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   570
      ScaleHeight     =   2085
      ScaleWidth      =   2865
      TabIndex        =   14
      Top             =   1590
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtCreditMemo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         MaxLength       =   6
         TabIndex        =   15
         ToolTipText     =   "Input Credit Memo Serial Number"
         Top             =   630
         Width           =   2415
      End
      Begin wizButton.cmd cmdPrintCreditMemo 
         Height          =   435
         Left            =   450
         TabIndex        =   16
         Top             =   1140
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         TX              =   "&Print"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print.frx":0FE6
      End
      Begin wizButton.cmd cmd5 
         Height          =   435
         Left            =   1530
         TabIndex        =   17
         Top             =   1140
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         TX              =   "&Back"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Print.frx":1002
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   2985
         _Version        =   655364
         _ExtentX        =   5265
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Credit Memo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         Alignment       =   1
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CREDIT MEMO NO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmSMIS_Report_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCustomer                                         As ADODB.Recordset
Dim rsPurchAgree                                       As ADODB.Recordset
Public GM                                              As String
Public IGNKEYNO                                        As String
Public VI_NO                                           As String

Private Sub cmd1_Click()

    On Error GoTo ErrorCode

    Screen.MousePointer = 11
    rptPrint.Reset
    ':::::::::::NEVER REMOVE THESE FORMULA::NOT APPLICABLE TO HAI,HBK,HGC,HMH,HSB,HAS
    LoadSignatories ("GATE PASS")
    '::::::CHECKED BY::::::
    rptPrint.Formulas(0) = "CHECKEDBY='" & Null2String(CheckedBy) & "'"
    rptPrint.Formulas(1) = "CHECKEDBYDESIG='" & Null2String(CheckedByDesig) & "'"
    '::::::APPROVED BY::::::
    rptPrint.Formulas(2) = "APPROVEDBY='" & Null2String(ApprovedBy) & "'"
    rptPrint.Formulas(3) = "SALESAPPROVEDDESIG='" & Null2String(SalesApprovedDesig) & "'"
    '::::::PREPARED BY::::::
    rptPrint.Formulas(4) = "PREPAREDBY='" & Null2String(PreparedBy) & "'"
    rptPrint.Formulas(5) = "PREPAREDBY='" & Null2String(PreparedByDesig) & "'"
    '::::::DELIVERD BY::::::
    rptPrint.Formulas(6) = "DELIVEREDBY='" & Null2String(DeliveredBy) & "'"
    rptPrint.Formulas(7) = "DELIVEREDBYDESIG='" & Null2String(DeliveredByDesig) & "'"
    '::::::GENERAL MANAGER::::::
    rptPrint.Formulas(8) = "GENERALMANAGER='" & Null2String(GeneralManager) & "'"
    rptPrint.Formulas(9) = "GENERALMANAGERDESIG='" & Null2String(GeneralManagerDesig) & "'"
    ':::::: SALES DISPATCHER ::::::
    rptPrint.Formulas(10) = "SALESDISPATCHER='" & Null2String(SalesDispatcher) & "'"
    rptPrint.Formulas(11) = "SALESDISPATCHERDESIG='" & Null2String(SalesDispatcherDesig) & "'"
    ':::::: FINANICING HEADS::::::
    rptPrint.Formulas(10) = "FINANCINGMANAGER='" & Null2String(FinancingManager) & "'"
    rptPrint.Formulas(11) = "FINANCINGMANAGERDESIG='" & Null2String(FinancingManagerDesig) & "'"
    
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "GatePass.rpt", "{SMIS_SalesOrder.VI_NO}='" & VI_NO & "'", DMIS_REPORT_Connection, 1
   
    '**************************
    NEW_LogAudit "V", "VEHICLE INVOICING", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "GatePass: INVOICE NO:" & VI_NO, "", ""
    '**************************
    Screen.MousePointer = 0

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmd4_Click()
    PRINTWARRANTYEXCEL IGNKEYNO
    '**************************
    NEW_LogAudit "V", "WARRANTY", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "Vehicle Invoicing: IGNEKY NO:" & IGNKEYNO, "", ""
    '**************************
End Sub

Private Sub cmd5_Click()
    picMain.Visible = True
    picCreditMemo.Visible = False
End Sub

Private Sub cmd6_Click()
    frmSMIS_Files_Signatories.Show 1
End Sub

Private Sub cmd7_Click()
    picMain.Visible = True
    picDebitMemo.Visible = False
End Sub

Private Sub cmdClearance_Click()
    rptPrint.Reset
    Screen.MousePointer = 11
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "clearance.rpt", "{customer.cuscde} = '" & CUSCODE & "' AND {PurchAgree.ProdNo} = '" & PRODUCTNO & "'", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdCreditMemo_Click()
    On Error GoTo ErrorCode
    Screen.MousePointer = 11
    rptPrint.Reset
    Dim temprs                                         As ADODB.Recordset
    rptPrint.Reset
    Set temprs = gconDMIS.Execute("select Count(*) from SMIS_MrrInv_Detail where  IgnKeyNo='" & IGNKEYNO & "'")
    If temprs(0).Value = 0 And rsPurchAgree!DISCOUNT = 0 Then
        MsgBox " There are No Record For This Transaction", vbInformation
        Exit Sub
    End If
    Dim RSCREDIT                                       As ADODB.Recordset
    Set RSCREDIT = gconDMIS.Execute("select creditmemo from smis_salesorder where VI_NO='" & VI_NO & "'")
    If IsNull(RSCREDIT("CREDITMEMO").Value) = True Then
        txtCreditMemo = (GenerateCode("SMIS_SALESORDER", "CREDITMEMO", "000000"))
    Else
        txtCreditMemo = RSCREDIT("CREDITMEMO").Value
    End If
    picCreditMemo.Visible = True
    picDebitMemo.Visible = False
    picMain.Visible = False
    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdDebitMemo_Click()

    On Error GoTo ErrorCode

    Screen.MousePointer = 11
    rptPrint.Reset
    Dim RSCREDIT                                       As ADODB.Recordset
    Set RSCREDIT = gconDMIS.Execute("SELECT DEBITMEMO FROM SMIS_SALESORDER WHERE VI_NO='" & VI_NO & "'")
    If IsNull(RSCREDIT("DEBITMEMO").Value) = True Then
        txtDebitMemo = (GenerateCode("SMIS_SALESORDER", "DEBITMEMO", "000000"))
    Else
        txtDebitMemo = RSCREDIT("DEBITMEMO").Value
    End If
    Set RSCREDIT = Nothing

    picDebitMemo.Visible = True
    picCreditMemo.Visible = False
    picMain.Visible = False
    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdDR_Click()
    Screen.MousePointer = 11
    rptPrint.Reset
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "dealers.rpt", "{customer.code} = '" & CUSCODE & "' AND {purchagree.prodno} = '" & PRODUCTNO & "'", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdjob_Click()
    On Error GoTo ErrorCode
    Screen.MousePointer = 11
    rptPrint.Reset
    ':::::::::::NEVER REMOVE THESE FORMULA:::::::::::'NOT APPLICABLE TO HAI,HBK,HGC,HMH,HSB,HAS
    LoadSignatories ("JOB REQUEST FORM")
    '::::::CHECKED BY::::::
    rptPrint.Formulas(0) = "CHECKEDBY='" & Null2String(CheckedBy) & "'"
    rptPrint.Formulas(1) = "CHECKEDBYDESIG='" & Null2String(CheckedByDesig) & "'"
    '::::::APPROVED BY::::::
    rptPrint.Formulas(2) = "APPROVEDBY='" & Null2String(ApprovedBy) & "'"
    rptPrint.Formulas(3) = "SALESAPPROVEDDESIG='" & Null2String(SalesApprovedDesig) & "'"
    '::::::PREPARED BY::::::
    rptPrint.Formulas(4) = "PREPAREDBY='" & Null2String(PreparedBy) & "'"
    rptPrint.Formulas(5) = "PreparedByDesig='" & Null2String(PreparedByDesig) & "'"
    '::::::DELIVERD BY::::::
    rptPrint.Formulas(6) = "DELIVEREDBY='" & Null2String(DeliveredBy) & "'"
    rptPrint.Formulas(7) = "DELIVEREDBYDESIG='" & Null2String(DeliveredByDesig) & "'"
    '::::::GENERAL MANAGER::::::
    rptPrint.Formulas(8) = "GENERALMANAGER='" & Null2String(GeneralManager) & "'"
    rptPrint.Formulas(9) = "GENERALMANAGERDESIG='" & Null2String(GeneralManagerDesig) & "'"
    ':::::: SALES DISPATCHER ::::::
    rptPrint.Formulas(10) = "SALESDISPATCHER='" & Null2String(SalesDispatcher) & "'"
    rptPrint.Formulas(11) = "SALESDISPATCHERDESIG='" & Null2String(SalesDispatcherDesig) & "'"
    ':::::: FINANICING HEADS::::::
    rptPrint.Formulas(10) = "FINANCINGMANAGER='" & Null2String(FinancingManager) & "'"
    rptPrint.Formulas(11) = "FINANCINGMANAGERDESIG='" & Null2String(FinancingManagerDesig) & "'"
    ':::::: COMPANY HEADS ::::::
    rptPrint.Formulas(12) = "Company_Name = '" & COMPANY_NAME & "'"
    rptPrint.Formulas(13) = "Company_Address = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "JOBREQUEST.RPT", "{PURCHAGREE.VI_NO} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
    '**************************
    NEW_LogAudit "V", "VEHICLE INVOICING", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "Job Request Form: INVOICE NO:" & VI_NO, "", ""
    '**************************
    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrintCreditMemo_Click()
    On Error GoTo ErrorCode
    Screen.MousePointer = 11
    ''''''
    Dim lng                                            As Integer
    lng = gconDMIS.Execute("select Count(*) from SMIS_SALESORDER WHERE CREDITMEMO=" & N2Str2Null(txtCreditMemo)).Fields(0).Value
    If lng >= 1 And UCase(Null2String(rsPurchAgree!CREDITMEMO)) <> UCase(txtCreditMemo) And Null2String(rsPurchAgree!CREDITMEMO) <> "" Then
        MessagePop RecSaveWarning, "Duplicate Record", "Credit Memo Number Already Exist"
        Screen.MousePointer = 0
        Exit Sub
    End If
    gconDMIS.Execute ("UPDATE SMIS_SALESORDER SET CREDITMEMO=" & N2Str2Null(txtCreditMemo) & " WHERE VI_NO='" & VI_NO & "'")
    rsRefresh
    rptPrint.Reset
    rptPrint.WindowShowPrintBtn = True
    ':::::::::::NEVER REMOVE THESE FORMULA::NOT APPLICABLE TO HAI,HBK,HGC,HMH,HSB,HAS
    LoadSignatories ("CREDIT MEMO")
    '::::::CHECKED BY::::::
    rptPrint.Formulas(0) = "CHECKEDBY='" & Null2String(CheckedBy) & "'"
    rptPrint.Formulas(1) = "CHECKEDBYDESIG='" & Null2String(CheckedByDesig) & "'"
    '::::::APPROVED BY::::::
    rptPrint.Formulas(2) = "APPROVEDBY='" & Null2String(ApprovedBy) & "'"
    rptPrint.Formulas(3) = "SALESAPPROVEDDESIG='" & Null2String(SalesApprovedDesig) & "'"
    '::::::PREPARED BY::::::
    rptPrint.Formulas(4) = "PREPAREDBY='" & Null2String(PreparedBy) & "'"
    rptPrint.Formulas(5) = "PreparedByDesig='" & Null2String(PreparedByDesig) & "'"
    '::::::DELIVERD BY::::::
    rptPrint.Formulas(6) = "DELIVEREDBY='" & Null2String(DeliveredBy) & "'"
    rptPrint.Formulas(7) = "DELIVEREDBYDESIG='" & Null2String(DeliveredByDesig) & "'"
    '::::::GENERAL MANAGER::::::
    rptPrint.Formulas(8) = "GENERALMANAGER='" & Null2String(GeneralManager) & "'"
    rptPrint.Formulas(9) = "GENERALMANAGERDESIG='" & Null2String(GeneralManagerDesig) & "'"
    ':::::: SALES DISPATCHER ::::::
    rptPrint.Formulas(10) = "SALESDISPATCHER='" & Null2String(SalesDispatcher) & "'"
    rptPrint.Formulas(11) = "SALESDISPATCHERDESIG='" & Null2String(SalesDispatcherDesig) & "'"
    ':::::: FINANICING HEADS::::::
    rptPrint.Formulas(10) = "FINANCINGMANAGER='" & Null2String(FinancingManager) & "'"
    rptPrint.Formulas(11) = "FINANCINGMANAGERDESIG='" & Null2String(FinancingManagerDesig) & "'"
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "CREDITMEMO.rpt", "{Purchagree.ignkey_no} = '" & IGNKEYNO & "'", DMIS_REPORT_Connection, 1
    '**************************
    NEW_LogAudit "V", "VEHICLE INVOICING", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "Credit Memo: INVOICE NO:" & VI_NO, "", ""
    '**************************
    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrintDebitmemo_Click()
    On Error GoTo ErrorCode
    Screen.MousePointer = 11
    Dim lng                                            As Integer
    lng = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_SALESORDER WHERE DEBITMEMO=" & N2Str2Null(txtDebitMemo)).Fields(0).Value
    If lng >= 1 And UCase(Null2String(rsPurchAgree!DEBITMEMO)) <> UCase(txtDebitMemo) And Null2String(rsPurchAgree!DEBITMEMO) <> "" Then
        MessagePop RecSaveWarning, "DUPLICATE RECORD", "DEBIT MEMO NUMBER ALREADY EXIST"
        Screen.MousePointer = 0
        Exit Sub
    End If
    gconDMIS.Execute ("UPDATE SMIS_SALESORDER SET DEBITMEMO=" & N2Str2Null(txtDebitMemo) & " WHERE VI_NO='" & VI_NO & "'")
    rptPrint.Reset
    rptPrint.WindowShowPrintBtn = True
    ':::::::::::NEVER REMOVE THESE FORMULA::NOT APPLICABLE TO HAI,HBK,HGC,HMH,HSB,HAS
    LoadSignatories ("DEBIT MEMO")
    '::::::CHECKED BY::::::
    rptPrint.Formulas(0) = "CHECKEDBY='" & Null2String(CheckedBy) & "'"
    rptPrint.Formulas(1) = "CHECKEDBYDESIG='" & Null2String(CheckedByDesig) & "'"
    '::::::APPROVED BY::::::
    rptPrint.Formulas(2) = "APPROVEDBY='" & Null2String(ApprovedBy) & "'"
    rptPrint.Formulas(3) = "SALESAPPROVEDDESIG='" & Null2String(SalesApprovedDesig) & "'"
    '::::::PREPARED BY::::::
    rptPrint.Formulas(4) = "PREPAREDBY='" & Null2String(PreparedBy) & "'"
    rptPrint.Formulas(5) = "PreparedByDesig='" & Null2String(PreparedByDesig) & "'"
    '::::::DELIVERD BY::::::
    rptPrint.Formulas(6) = "DELIVEREDBY='" & Null2String(DeliveredBy) & "'"
    rptPrint.Formulas(7) = "DELIVEREDBYDESIG='" & Null2String(DeliveredByDesig) & "'"
    '::::::GENERAL MANAGER::::::
    rptPrint.Formulas(8) = "GENERALMANAGER='" & Null2String(GeneralManager) & "'"
    rptPrint.Formulas(9) = "GENERALMANAGERDESIG='" & Null2String(GeneralManagerDesig) & "'"
    ':::::: SALES DISPATCHER ::::::
    rptPrint.Formulas(10) = "SALESDISPATCHER='" & Null2String(SalesDispatcher) & "'"
    rptPrint.Formulas(11) = "SALESDISPATCHERDESIG='" & Null2String(SalesDispatcherDesig) & "'"
    ':::::: FINANICING HEADS::::::
    rptPrint.Formulas(10) = "FINANCINGMANAGER='" & Null2String(FinancingManager) & "'"
    rptPrint.Formulas(11) = "FINANCINGMANAGERDESIG='" & Null2String(FinancingManagerDesig) & "'"
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "DEBITMEMO.RPT", "{PurchAgree.VI_NO} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
    '**************************
    NEW_LogAudit "V", "VEHICLE INVOICING", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "Debit Memo: INVOICE NO:" & VI_NO, "", ""
    '**************************
    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdReleaseOrder_Click()
    On Error GoTo ErrorCode
    Screen.MousePointer = 11
    rptPrint.Reset
    LoadSignatories ("RELEASE ORDER")
    ':::::::::::NEVER REMOVE THESE FORMULA::NOT APPLICABLE TO HAI,HBK,HGC,HMH,HSB,HAS
    '::::::CHECKED BY::::::
    rptPrint.Formulas(0) = "CHECKEDBY='" & Null2String(CheckedBy) & "'"
    rptPrint.Formulas(1) = "CHECKEDBYDESIG='" & Null2String(CheckedByDesig) & "'"
    '::::::APPROVED BY::::::
    rptPrint.Formulas(2) = "APPROVEDBY='" & Null2String(ApprovedBy) & "'"
    rptPrint.Formulas(3) = "SALESAPPROVEDDESIG='" & Null2String(SalesApprovedDesig) & "'"
    '::::::PREPARED BY::::::
    rptPrint.Formulas(4) = "PREPAREDBY='" & Null2String(PreparedBy) & "'"
    rptPrint.Formulas(5) = "PreparedByDesig='" & Null2String(PreparedByDesig) & "'"
    '::::::DELIVERD BY::::::
    rptPrint.Formulas(6) = "DELIVEREDBY='" & Null2String(DeliveredBy) & "'"
    rptPrint.Formulas(7) = "DELIVEREDBYDESIG='" & Null2String(DeliveredByDesig) & "'"
    '::::::GENERAL MANAGER::::::
    rptPrint.Formulas(8) = "GENERALMANAGER='" & Null2String(GeneralManager) & "'"
    rptPrint.Formulas(9) = "GENERALMANAGERDESIG='" & Null2String(GeneralManagerDesig) & "'"
    ':::::: SALES DISPATCHER ::::::
    rptPrint.Formulas(10) = "SALESDISPATCHER='" & Null2String(SalesDispatcher) & "'"
    rptPrint.Formulas(11) = "SALESDISPATCHERDESIG='" & Null2String(SalesDispatcherDesig) & "'"
    ':::::: FINANICING HEADS::::::
    rptPrint.Formulas(10) = "FINANCINGMANAGER='" & Null2String(FinancingManager) & "'"
    rptPrint.Formulas(11) = "FINANCINGMANAGERDESIG='" & Null2String(FinancingManagerDesig) & "'"
    ':::::: COMPANY HEADS ::::::
    rptPrint.Formulas(12) = "COMPANYNAME = '" & COMPANY_NAME & "'"
    rptPrint.Formulas(13) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "ReleaseOrder.rpt", "{purchagree.VI_NO} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
    '**************************
    NEW_LogAudit "V", "RELEASED ORDER", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "Released Order: INVOICE NO:" & VI_NO, "", ""
    '**************************
    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdTransaction_Click()
    On Error GoTo ErrorCode
    Screen.MousePointer = 11
    rptPrint.Reset
    ':::::::::::NEVER REMOVE THESE FORMULA:::::::::::'NOT APPLICABLE TO HAI,HBK,HGC,HMH,HSB,HAS
    LoadSignatories ("TRANSACTION SLIP")
    '::::::CHECKED BY::::::
    rptPrint.Formulas(0) = "CHECKEDBY='" & Null2String(CheckedBy) & "'"
    rptPrint.Formulas(1) = "CHECKEDBYDESIG='" & Null2String(CheckedByDesig) & "'"
    '::::::APPROVED BY::::::
    rptPrint.Formulas(2) = "APPROVEDBY='" & Null2String(ApprovedBy) & "'"
    rptPrint.Formulas(3) = "SALESAPPROVEDDESIG='" & Null2String(SalesApprovedDesig) & "'"
    '::::::PREPARED BY::::::
    rptPrint.Formulas(4) = "PREPAREDBY='" & Null2String(PreparedBy) & "'"
    rptPrint.Formulas(5) = "PreparedByDesig='" & Null2String(PreparedByDesig) & "'"
    '::::::DELIVERD BY::::::
    rptPrint.Formulas(6) = "DELIVEREDBY='" & Null2String(DeliveredBy) & "'"
    rptPrint.Formulas(7) = "DELIVEREDBYDESIG='" & Null2String(DeliveredByDesig) & "'"
    '::::::GENERAL MANAGER::::::
    rptPrint.Formulas(8) = "GENERALMANAGER='" & Null2String(GeneralManager) & "'"
    rptPrint.Formulas(9) = "GENERALMANAGERDESIG='" & Null2String(GeneralManagerDesig) & "'"
    ':::::: SALES DISPATCHER ::::::
    rptPrint.Formulas(10) = "SALESDISPATCHER='" & Null2String(SalesDispatcher) & "'"
    rptPrint.Formulas(11) = "SALESDISPATCHERDESIG='" & Null2String(SalesDispatcherDesig) & "'"
    ':::::: FINANICING HEADS::::::
    rptPrint.Formulas(10) = "FINANCINGMANAGER='" & Null2String(FinancingManager) & "'"
    rptPrint.Formulas(11) = "FINANCINGMANAGERDESIG='" & Null2String(FinancingManagerDesig) & "'"
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "Transactionslip.rpt", "{purchagree.VI_NO} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
    '**************************
    NEW_LogAudit "V", "VEHICLE INVOICING", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "Transaction Slip: INVOICE NO:" & VI_NO, "", ""
    '**************************
    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdVDR_Click()
    On Error GoTo ErrorCode
    Screen.MousePointer = 11
    rptPrint.Reset
    ':::::::::::NEVER REMOVE THESE FORMULA::NOT APPLICABLE TO HAI,HBK,HGC,HMH,HSB,HAS
    LoadSignatories ("DELIVERY REPORT")
    '::::::CHECKED BY::::::
    rptPrint.Formulas(0) = "CHECKEDBY='" & Null2String(CheckedBy) & "'"
    rptPrint.Formulas(1) = "CHECKEDBYDESIG='" & Null2String(CheckedByDesig) & "'"
    '::::::APPROVED BY::::::
    rptPrint.Formulas(2) = "APPROVEDBY='" & Null2String(ApprovedBy) & "'"
    rptPrint.Formulas(3) = "SALESAPPROVEDDESIG='" & Null2String(SalesApprovedDesig) & "'"
    '::::::PREPARED BY::::::
    rptPrint.Formulas(4) = "PreparedBy='" & Null2String(PreparedBy) & "'"
    rptPrint.Formulas(5) = "PreparedByDesig='" & Null2String(PreparedByDesig) & "'"
    '::::::DELIVERD BY::::::
    rptPrint.Formulas(6) = "DELIVEREDBY='" & Null2String(DeliveredBy) & "'"
    rptPrint.Formulas(7) = "DELIVEREDBYDESIG='" & Null2String(DeliveredByDesig) & "'"
    '::::::GENERAL MANAGER::::::
    rptPrint.Formulas(8) = "GENERALMANAGER='" & Null2String(GeneralManager) & "'"
    rptPrint.Formulas(9) = "GENERALMANAGERDESIG='" & Null2String(GeneralManagerDesig) & "'"
    ':::::: SALES DISPATCHER ::::::
    rptPrint.Formulas(10) = "SALESDISPATCHER='" & Null2String(SalesDispatcher) & "'"
    rptPrint.Formulas(11) = "SALESDISPATCHERDESIG='" & Null2String(SalesDispatcherDesig) & "'"
    ':::::: FINANICING HEADS::::::
    rptPrint.Formulas(10) = "FINANCINGMANAGER='" & Null2String(FinancingManager) & "'"
    rptPrint.Formulas(11) = "FINANCINGMANAGERDESIG='" & Null2String(FinancingManagerDesig) & "'"
'updated by:    IEBV 0942010_0245pm
'description:   To Desplay free beeies on report for the HPC Dealer
'-----------------------------------------------------------------------------------------------------
   If COMPANY_CODE = "HPC" Then
    Dim rsfree                              As New ADODB.Recordset
    Dim desfree                             As String
    Dim ctr                                 As Integer
    desfree = ""
    Set rsfree = New ADODB.Recordset
    Set rsfree = gconDMIS.Execute("select description from SMIS_MRRINV_DETAIL where IgnKeyNo = '" & IGNKEYNO & "'")
        If Not rsfree.EOF And Not rsfree.BOF Then
            rsfree.MoveFirst
            ctr = 1
            Do While Not rsfree.EOF
                If ctr = 1 Then
                    desfree = desfree + " " + rsfree.Fields(0).Value
                    ctr = ctr + 1
                Else
                    desfree = desfree + ", " + rsfree.Fields(0).Value
                    ctr = ctr + 1
                End If
                rsfree.MoveNext
            Loop
            rptPrint.Formulas(12) = "FREE = '" & Null2String(desfree) & "'"
        End If
    End If
'-----------------------------------------------------------------------------------------------------
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "vdr.rpt", "{purchagree.VI_NO} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
    '**************************
    NEW_LogAudit "V", "VEHICLE INVOICING", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "Vehicile Invoicing: VDR NO:" & Null2String(frmSMIS_Trans_VehicleInvoice.txtRelease_VDR.Text), "", ""
    '**************************
    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdVI_Click()

    On Error GoTo ErrorCode

    Screen.MousePointer = 11
    rptPrint.Reset
    ':::::::::::NEVER REMOVE THESE FORMULA::NOT APPLICABLE TO HAI,HBK,HGC,HMH,HSB,HAS
    LoadSignatories ("SALES INVOICE")
    '::::::CHECKED BY::::::
    rptPrint.Formulas(0) = "CHECKEDBY='" & Null2String(CheckedBy) & "'"
    rptPrint.Formulas(1) = "CHECKEDBYDESIG='" & Null2String(CheckedByDesig) & "'"
    '::::::APPROVED BY::::::
    rptPrint.Formulas(2) = "APPROVEDBY='" & Null2String(ApprovedBy) & "'"
    rptPrint.Formulas(3) = "SALESAPPROVEDDESIG='" & Null2String(SalesApprovedDesig) & "'"
    '::::::PREPARED BY::::::
    rptPrint.Formulas(4) = "PREPAREDBY='" & Null2String(PreparedBy) & "'"
    rptPrint.Formulas(5) = "PreparedByDesig='" & Null2String(PreparedByDesig) & "'"
    '::::::DELIVERD BY::::::
    rptPrint.Formulas(6) = "DELIVEREDBY='" & Null2String(DeliveredBy) & "'"
    rptPrint.Formulas(7) = "DELIVEREDBYDESIG='" & Null2String(DeliveredByDesig) & "'"
    '::::::GENERAL MANAGER::::::
    rptPrint.Formulas(8) = "GENERALMANAGER='" & Null2String(GeneralManager) & "'"
    rptPrint.Formulas(9) = "GENERALMANAGERDESIG='" & Null2String(GeneralManagerDesig) & "'"
    ':::::: SALES DISPATCHER ::::::
    rptPrint.Formulas(10) = "SALESDISPATCHER='" & Null2String(SalesDispatcher) & "'"
    rptPrint.Formulas(11) = "SALESDISPATCHERDESIG='" & Null2String(SalesDispatcherDesig) & "'"
    ':::::: FINANICING HEADS::::::
    rptPrint.Formulas(10) = "FINANCINGMANAGER='" & Null2String(FinancingManager) & "'"
    rptPrint.Formulas(11) = "FINANCINGMANAGERDESIG='" & Null2String(FinancingManagerDesig) & "'"
'updated by:    IEBV 0942010_0245pm
'description:   To Desplay free beeies on report for the HPC Dealer
'-----------------------------------------------------------------------------------------------------
    If COMPANY_CODE = "HPC" Then
    Dim rsfree                              As New ADODB.Recordset
    Dim desfree                             As String
    Dim ctr                                 As Integer
    desfree = ""
    Set rsfree = New ADODB.Recordset
    Set rsfree = gconDMIS.Execute("select description from SMIS_MRRINV_DETAIL where IgnKeyNo = '" & IGNKEYNO & "'")
        If Not rsfree.EOF And Not rsfree.BOF Then
            rsfree.MoveFirst
            ctr = 1
            Do While Not rsfree.EOF
                If ctr = 1 Then
                    desfree = desfree + " " + rsfree.Fields(0).Value
                    ctr = ctr + 1
                Else
                    desfree = desfree + ", " + rsfree.Fields(0).Value
                    ctr = ctr + 1
                End If
                rsfree.MoveNext
            Loop
            rptPrint.Formulas(12) = "FREE = '" & Null2String(desfree) & "'"
        End If
    End If
'-----------------------------------------------------------------------------------------------------

    If Null2String(rsPurchAgree!TERM) = "COD" Or Null2String(rsPurchAgree!TERM) = "CPO" Then
        Screen.MousePointer = 11
        If COMPANY_CODE = "HSR" Then
            rptPrint.Formulas(0) = "PreparedBy='" & Null2String(PreparedBy) & "'"
            rptPrint.Formulas(1) = "CheckedBy='" & Null2String(CheckedBy) & "'"
            rptPrint.Formulas(2) = "ApprovedBy='" & Null2String(ApprovedBy) & "'"
            rptPrint.Formulas(3) = "GM='" & Null2String(GeneralManager) & "'"
            PrintSQLReport rptPrint, SMIS_REPORT_PATH & "vi_compo.rpt", "{purchagree.VI_NO} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
        Else
            PrintSQLReport rptPrint, SMIS_REPORT_PATH & "VI.RPT", "{PURCHAGREE.VI_NO} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
        End If
        Screen.MousePointer = 0
        
    ElseIf Null2String(rsPurchAgree!TERM) = "BPO" Then
            Screen.MousePointer = 11
            rptPrint.Formulas(0) = "PreparedBy='" & Null2String(PreparedBy) & "'"
            rptPrint.Formulas(1) = "CheckedBy='" & Null2String(CheckedBy) & "'"
            rptPrint.Formulas(2) = "ApprovedBy='" & Null2String(ApprovedBy) & "'"
            rptPrint.Formulas(3) = "GM='" & Null2String(GeneralManager) & "'"
            PrintSQLReport rptPrint, SMIS_REPORT_PATH & "VI_FIN_BPO.rpt", "{purchagree.VI_NO} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
    Else
        Screen.MousePointer = 11
        PrintSQLReport rptPrint, SMIS_REPORT_PATH & "VI_FIN.RPT", "{PURCHAGREE.VI_NO} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    End If
     
    '**************************
    NEW_LogAudit "V", "VEHICLE INVOICING", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "INVOICE NO:" & VI_NO, "", ""
    '**************************
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "select * from ALL_CUSTOMER_TABLE where CUSCDE = '" & CUSCODE & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If rsCustomer.BOF And rsCustomer.EOF Then
        MsgSpeechBox "Error Encountered! Empty Customer Record!"
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GM = ""
    IGNKEYNO = ""
    VI_NO = ""
End Sub

Sub PRINTWARRANTYEXCEL(xYear)
    If Len(Dir(App.Path & "\warranty.xls")) = 0 Then
        MsgBox "Could Not located Warranty.Xls File", vbInformation
        Exit Sub
    End If
    On Error GoTo ErrorCode
    Dim xlApp                                          As Excel.Application
    Dim xlBook                                         As Excel.Workbook
    Dim xlSheet                                        As Excel.Worksheet
    Set xlApp = New Excel.Application

    Set xlBook = xlApp.Workbooks.Open(App.Path & "\warranty.xls")
    Set xlSheet = xlBook.Worksheets(1)
    Dim rsModel                                        As ADODB.Recordset
    Dim vmodel                                         As String
    Dim i                                              As Integer
    Dim j                                              As Integer
    Dim rsCountProspect                                As ADODB.Recordset
    Set rsModel = gconDMIS.Execute("Select * from SMIS_SALESORDER where VI_NO='" & VI_NO & "'")
    If Not rsModel.EOF Or Not rsModel.BOF Then
        xlSheet.Cells(1, 1) = Null2String(rsModel("CUSTNAME"))
        xlSheet.Cells(3, 1) = Null2String(rsModel("HOMEADDRESS"))
        xlSheet.Cells(9, 1) = Null2String(rsModel("INVOICEDDATE"))
        xlSheet.Cells(11, 1) = Null2String(rsModel("MODELDESCRIPTION"))
        xlSheet.Cells(13, 1) = Null2String(rsModel("IGNKEY_NO"))
        xlSheet.Cells(19, 2) = Null2String(rsModel("VINO"))
        xlSheet.Cells(22, 1) = Null2String(rsModel("ENGINENO"))
        xlSheet.Cells(22, 2) = Null2String(rsModel("COLOR"))
        xlSheet.Cells(22, 3) = Null2String(rsModel("CUSTNAME"))
        xlSheet.Cells(22, 4) = Null2String(rsModel("DATERELEASED"))
        xlApp.Visible = True
        Set xlApp = Nothing
    End If
    Exit Sub
ErrorCode:
    MsgBox Err.Description
    Err.Clear
End Sub

Sub rsRefresh()
    Set rsPurchAgree = New ADODB.Recordset
    rsPurchAgree.Open "select * from SMIS_SALESORDER where code = '" & CUSCODE & "' AND IGNKEY_NO ='" & IGNKEYNO & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub txtCreditMemo_LostFocus()
    txtCreditMemo = Format(txtCreditMemo, "000000")
End Sub


VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmSMIS_Trans_VehicleInvoicePrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing Process..."
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
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
   Icon            =   "VehicleInvoicePrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4380
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   5325
      Left            =   0
      ScaleHeight     =   5325
      ScaleWidth      =   6360
      TabIndex        =   0
      Top             =   0
      Width           =   6360
      Begin Crystal.CrystalReport rptPrint 
         Left            =   390
         Top             =   570
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin wizButton.cmd cmdCreditMemo 
         Height          =   435
         Left            =   180
         TabIndex        =   4
         ToolTipText     =   "Release Order"
         Top             =   1650
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
         MICON           =   "VehicleInvoicePrint.frx":0E42
      End
      Begin wizButton.cmd cmdDebitMemo 
         Height          =   435
         Left            =   180
         TabIndex        =   5
         ToolTipText     =   "PNP Motor Vehicle Clearance Application"
         Top             =   2145
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
         MICON           =   "VehicleInvoicePrint.frx":0E5E
      End
      Begin wizButton.cmd cmdVDR 
         Height          =   435
         Left            =   180
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
         MICON           =   "VehicleInvoicePrint.frx":0E7A
      End
      Begin wizButton.cmd cmdVI 
         Height          =   435
         Left            =   180
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
         MICON           =   "VehicleInvoicePrint.frx":0E96
      End
      Begin wizButton.cmd cmdExit 
         Height          =   435
         Left            =   180
         TabIndex        =   6
         ToolTipText     =   "Exit"
         Top             =   3630
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
         MICON           =   "VehicleInvoicePrint.frx":0EB2
      End
      Begin wizButton.cmd cmd1 
         Height          =   435
         Left            =   180
         TabIndex        =   3
         ToolTipText     =   "Gate Pass"
         Top             =   1155
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
         MICON           =   "VehicleInvoicePrint.frx":0ECE
      End
      Begin wizButton.cmd cmdReleaseOrder 
         Height          =   435
         Left            =   150
         TabIndex        =   7
         ToolTipText     =   "PNP Motor Vehicle Clearance Application"
         Top             =   2640
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
         MICON           =   "VehicleInvoicePrint.frx":0EEA
      End
      Begin wizButton.cmd cmdClearance 
         Height          =   435
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "PNP Motor Vehicle Clearance Application"
         Top             =   4020
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
         MICON           =   "VehicleInvoicePrint.frx":0F06
      End
      Begin wizButton.cmd cmdDR 
         Height          =   435
         Left            =   120
         TabIndex        =   9
         Top             =   4500
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
         MICON           =   "VehicleInvoicePrint.frx":0F22
      End
      Begin wizButton.cmd cmd4 
         Height          =   435
         Left            =   180
         TabIndex        =   23
         ToolTipText     =   "Exit"
         Top             =   3135
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
         MICON           =   "VehicleInvoicePrint.frx":0F3E
      End
   End
   Begin VB.PictureBox picDebitMemo 
      BorderStyle     =   0  'None
      Height          =   5325
      Left            =   0
      ScaleHeight     =   5325
      ScaleWidth      =   6360
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   6360
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
         Height          =   435
         Left            =   600
         MaxLength       =   6
         TabIndex        =   12
         ToolTipText     =   "Input Debit Memo Serial Number"
         Top             =   750
         Width           =   2445
      End
      Begin wizButton.cmd cmdPrintDebitmemo 
         Height          =   435
         Left            =   960
         TabIndex        =   13
         Top             =   1380
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
         MICON           =   "VehicleInvoicePrint.frx":0F5A
      End
      Begin wizButton.cmd cmd7 
         Height          =   435
         Left            =   2070
         TabIndex        =   14
         Top             =   1380
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
         MICON           =   "VehicleInvoicePrint.frx":0F76
      End
      Begin VB.Label Label3 
         Caption         =   "DEBIT MEMO #"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   630
         TabIndex        =   11
         Top             =   450
         Width           =   1995
      End
   End
   Begin VB.PictureBox picCreditMemo 
      BorderStyle     =   0  'None
      Height          =   5325
      Left            =   0
      ScaleHeight     =   5325
      ScaleWidth      =   6360
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   6360
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
         Height          =   465
         Left            =   630
         MaxLength       =   6
         TabIndex        =   20
         ToolTipText     =   "Input Credit Memo Serial Number"
         Top             =   630
         Width           =   2415
      End
      Begin wizButton.cmd cmdPrintCreditMemo 
         Height          =   435
         Left            =   960
         TabIndex        =   21
         Top             =   1260
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
         MICON           =   "VehicleInvoicePrint.frx":0F92
      End
      Begin wizButton.cmd cmd5 
         Height          =   435
         Left            =   2070
         TabIndex        =   22
         Top             =   1260
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
         MICON           =   "VehicleInvoicePrint.frx":0FAE
      End
      Begin VB.Label Label4 
         Caption         =   "CREDIT MEMO #"
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
         Left            =   630
         TabIndex        =   19
         Top             =   330
         Width           =   1995
      End
   End
   Begin VB.PictureBox picGatePass 
      BorderStyle     =   0  'None
      Height          =   5325
      Left            =   -30
      ScaleHeight     =   5325
      ScaleWidth      =   6360
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   6360
      Begin wizButton.cmd cmd2 
         Height          =   435
         Left            =   1620
         TabIndex        =   16
         Top             =   1680
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
         MICON           =   "VehicleInvoicePrint.frx":0FCA
      End
      Begin wizButton.cmd cmd3 
         Height          =   435
         Left            =   2730
         TabIndex        =   17
         Top             =   1680
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         TX              =   "&Cancel"
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
         MICON           =   "VehicleInvoicePrint.frx":0FE6
      End
   End
End
Attribute VB_Name = "frmSMIS_Trans_VehicleInvoicePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCustomer                                                        As ADODB.Recordset
Dim rsPurchAgree                                                      As ADODB.Recordset
Public GM                                                             As String
Public IGNKEYNO                                                       As String

Private Sub cmd1_Click()
    Screen.MousePointer = 11
    rptPrint.Reset
    picMain.Visible = False
    picGatePass.Visible = True
    Screen.MousePointer = 0
End Sub

Private Sub cmd2_Click()
    Screen.MousePointer = 11
    rptPrint.Reset
    LoadSignatories ("GATE PASS")
    'rptPrint.Formulas(0) = "GuardOnDuty=" & N2Str2Null(txtGatePassGuardOnDuty)
    'rptPrint.Formulas(1) = "TimeOut=" & N2Str2Null(txtGatePassTimeOut)
    'rptPrint.Formulas(0) = "FinancingManager=" & N2Str2Null(FinancingManager)
    
    rptPrint.Formulas(1) = "ApprovedBy=" & N2Str2Null(ApprovedBy)
    rptPrint.Formulas(2) = "ReleasedBy=" & N2Str2Null(FinancingManager)
    rptPrint.Formulas(3) = "GM=" & N2Str2Null(GeneralManager)
    
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "GatePass.rpt", "{SMIS_SalesOrder.IGNKEY_NO}='" & IGNKEYNO & "'", DMIS_REPORT_Connection, 1
    LogAudit "V", "GATE PASS ", "VI_NO" & Null2String(rsPurchAgree!VI_NO)
    cmd3_Click
    Screen.MousePointer = 0
End Sub

Private Sub cmd3_Click()
    picMain.Visible = True
    picGatePass.Visible = False
End Sub

Private Sub cmd4_Click()
    frmSMIS_FILES_Signatories.Show 1
End Sub

Private Sub cmdClearance_Click()
    rptPrint.Reset
    Screen.MousePointer = 11
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "clearance.rpt", "{customer.cuscde} = '" & CusCode & "' AND {PurchAgree.ProdNo} = '" & PRODUCTNO & "'", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmd5_Click()
    picMain.Visible = True
    picCreditMemo.Visible = False
End Sub

Private Sub cmd7_Click()
    picMain.Visible = True
    picDebitMemo.Visible = False
End Sub

Private Sub cmdCreditMemo_Click()
    Dim temprs                                                        As ADODB.Recordset
    rptPrint.Reset
    Set temprs = gconDMIS.Execute("select Count(*) from SMIS_MrrInv_Detail where  IgnKeyNo='" & IGNKEYNO & "'")
    If temprs(0).Value = 0 And rsPurchAgree!DISCOUNT = 0 Then
        MsgBox " There are No Record For This Transaction", vbInformation
        Exit Sub
    End If
    Dim RSCREDIT                                                      As ADODB.Recordset
    Set RSCREDIT = gconDMIS.Execute("select creditmemo from smis_salesorder where ignkey_no='" & IGNKEYNO & "'")
    If IsNull(RSCREDIT("CREDITMEMO").Value) = True Then
        txtCreditMemo = (GenerateCode("SMIS_SALESORDER", "CREDITMEMO", "000000"))
    Else
        txtCreditMemo = RSCREDIT("CREDITMEMO").Value
    End If
    picCreditMemo.Visible = True
    picMain.Visible = False
End Sub
Private Sub cmdDebitMemo_Click()
    rptPrint.Reset
    ' If Null2String(rsPurchAgree!Term) = "F" Then
    Dim RSCREDIT                                                      As ADODB.Recordset
    Set RSCREDIT = gconDMIS.Execute("SELECT DEBITMEMO FROM SMIS_SALESORDER WHERE IGNKEY_NO='" & IGNKEYNO & "'")
    If IsNull(RSCREDIT("DEBITMEMO").Value) = True Then
        txtDebitMemo = (GenerateCode("SMIS_SALESORDER", "DEBITMEMO", "000000"))
    Else
        txtDebitMemo = RSCREDIT("DEBITMEMO").Value
    End If
    Set RSCREDIT = Nothing
    picDebitMemo.Visible = True
    picMain.Visible = False
    ' Else
    '     MsgBox " Not Applicable for this Type of Transaction .", vbExclamation
    '  End If
End Sub

Private Sub cmdDR_Click()
    Screen.MousePointer = 11
    rptPrint.Reset
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "dealers.rpt", "{customer.code} = '" & CusCode & "' AND {purchagree.prodno} = '" & PRODUCTNO & "'", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrintCreditMemo_Click()
    Screen.MousePointer = 11
    '''''''AXP063020071200
    Dim lng                                                           As Integer
    lng = gconDMIS.Execute("select Count(*) from SMIS_SALESORDER WHERE CREDITMEMO=" & N2Str2Null(txtCreditMemo)).Fields(0).Value
    If lng >= 1 And UCase(Null2String(rsPurchAgree!CREDITMEMO)) <> UCase(txtCreditMemo) Then
        MessagePop RecSaveWarning, "Duplicate Record", "Credit Memo Number Already Exist"
        Screen.MousePointer = 0
        Exit Sub
    End If
    gconDMIS.Execute ("UPDATE SMIS_SALESORDER SET CREDITMEMO=" & N2Str2Null(txtCreditMemo) & " WHERE IGNKEY_NO='" & IGNKEYNO & "'")
    rsRefresh
    rptPrint.WindowShowPrintBtn = True
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "CREDITMEMO.rpt", "{MRR.ignkey} = '" & IGNKEYNO & "'", DMIS_REPORT_Connection, 1
    LogAudit "V", "CREDIT MEMO", "VI NO#" & Null2String(rsPurchAgree!VI_NO)
    Screen.MousePointer = 0
End Sub

Private Sub cmdPrintDebitmemo_Click()
    Screen.MousePointer = 11
    '''''''AXP063020071200
    Dim lng                                                           As Integer
    lng = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_SALESORDER WHERE DEBITMEMO=" & N2Str2Null(txtDebitMemo)).Fields(0).Value
    If lng >= 1 And UCase(Null2String(rsPurchAgree!DEBITMEMO)) <> UCase(txtDebitMemo) Then
        MessagePop RecSaveWarning, "DUPLICATE RECORD", "DEBIT MEMO NUMBER ALREADY EXIST"
        Screen.MousePointer = 0
        Exit Sub
    End If
    gconDMIS.Execute ("UPDATE SMIS_SALESORDER SET DEBITMEMO=" & N2Str2Null(txtDebitMemo) & " WHERE IGNKEY_NO='" & IGNKEYNO & "'")
    rsRefresh
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "DEBITMEMO.RPT", "{PurchAgree.IGNKEY_NO} = '" & IGNKEYNO & "'", DMIS_REPORT_Connection, 1
    LogAudit "V", "DEBIT MEMO", "VI NO#" & Null2String(rsPurchAgree!VI_NO)
    Screen.MousePointer = 0
End Sub

Private Sub cmdReleaseOrder_Click()
    rptPrint.Reset
    'If Not rsCustomer.EOF And Not rsCustomer.BOF Then
    '    If Null2String(rsPurchAgree!Term) = "COD" Then
    Screen.MousePointer = 11
    rptPrint.Reset
    rptPrint.Formulas(0) = "CompanyName = '" & Company_name & "'"
    rptPrint.Formulas(1) = "CompanyAddress = '" & Company_Address & "'"
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "ReleaseOrder.rpt", "{purchagree.ignKey_no} = '" & IGNKEYNO & "'", DMIS_REPORT_Connection, 1
    LogAudit "V", "RELEASE ORDER", "CS#" & Null2String(rsPurchAgree!IGNKEY_NO)
    Screen.MousePointer = 0

    '     Else
    '         Screen.MousePointer = 11
    '         PrintSQLReport rptPrint, SMIS_REPORT_PATH & "releaseFI.rpt", "{customer.code} = '" & CusCode & "' AND {purchagree.prodno} = '" & PRODUCTNO & "'", DMIS_REPORT_Connection, 1
    '         Screen.MousePointer = 0
    '     End If
    'End If
End Sub
Private Sub cmdVDR_Click()
    rptPrint.Reset
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        Screen.MousePointer = 11
        LoadSignatories ("DELIVERY REPORT")
        rptPrint.Formulas(0) = "PreparedBy=" & N2Str2Null(PreparedBy)
        rptPrint.Formulas(1) = "CheckedBy=" & N2Str2Null(CheckedBy)
        rptPrint.Formulas(2) = "SalesApproved=" & N2Str2Null(ApprovedBy)
        rptPrint.Formulas(3) = "GM=" & N2Str2Null(GeneralManager)
        PrintSQLReport rptPrint, SMIS_REPORT_PATH & "vdr.rpt", "{customer.cuscde} = '" & CusCode & "' AND {purchagree.ignKey_no} = '" & IGNKEYNO & "'", DMIS_REPORT_Connection, 1
        LogAudit "V", "VDR", "VDR NO" & Null2String(rsPurchAgree!VDR_NO)
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record Found!"
    End If
End Sub

'Upating Code       : AXP-0707200712:44
'Upating Code       : AXP-0707200712:44
Private Sub cmdVI_Click()
    On Error GoTo ErrorCode:
    rptPrint.Reset
    LoadSignatories ("SALES INVOICE")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then

        If Null2String(rsPurchAgree!Term) = "COD" Then
            Screen.MousePointer = 11
            rptPrint.Formulas(0) = "GM=" & N2Str2Null(GeneralManager)
            PrintSQLReport rptPrint, SMIS_REPORT_PATH & "vi.rpt", "{customer.CUSCDE} = '" & CusCode & "' AND {purchagree.prodno} = '" & PRODUCTNO & "'", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
            
        Else
            Screen.MousePointer = 11
            rptPrint.Formulas(0) = "GM=" & N2Str2Null(GeneralManager)
            PrintSQLReport rptPrint, SMIS_REPORT_PATH & "vi.rpt", "{customer.CUSCDE} = '" & CusCode & "' AND {purchagree.prodno} = '" & PRODUCTNO & "'", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        End If
        LogAudit "V", "VI#", "VI:" & Null2String(rsPurchAgree!VI_NO)
        
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    rsRefresh
    'Dim RSSIGNATORIES
    '  Set RSSIGNATORIES = New ADODB.Recordset
    '        RSSIGNATORIES.Open "select * from SMIS_Signatories", gconDMIS, adOpenForwardOnly, adLockReadOnly
    '        If Not RSSIGNATORIES.EOF And Not RSSIGNATORIES.BOF Then
    '            GM = Null2String(RSSIGNATORIES!GeneralManager)
    '        End If
    '    If Not rsPurchAgree.EOF And Not rsPurchAgree.BOF Then
    '
    '
    '        If Null2String(rsPurchAgree!DateReleased) = "" Then
    '            cmdCreditMemo.Enabled = False
    '            cmdVDR.Enabled = False
    '        Else
    '            cmdCreditMemo.Enabled = True
    '            cmdVDR.Enabled = True
    '        End If
    Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "select * from ALL_CUSTMASTER_SMIS where code = '" & CusCode & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If rsCustomer.BOF And rsCustomer.EOF Then
        MsgSpeechBox "Error Encountered! Empty Customer Record!"
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GM = ""
    IGNKEYNO = ""
End Sub

Sub rsRefresh()
    Set rsPurchAgree = New ADODB.Recordset

    rsPurchAgree.Open "select * from SMIS_SALESORDER where code = '" & CusCode & "' AND IGNKEY_NO ='" & IGNKEYNO & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub txtCreditMemo_LostFocus()
    txtCreditMemo = Format(txtCreditMemo, "000000")
End Sub


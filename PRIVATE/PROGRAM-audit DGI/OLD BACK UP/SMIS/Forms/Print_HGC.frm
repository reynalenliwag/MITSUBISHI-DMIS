VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmSMIS_Report_Print_HGC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing Process..."
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4320
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
   Icon            =   "Print_HGC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   4320
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   5745
      Left            =   0
      ScaleHeight     =   5745
      ScaleWidth      =   6360
      TabIndex        =   0
      Top             =   0
      Width           =   6360
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
         MICON           =   "Print_HGC.frx":0E42
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
         MICON           =   "Print_HGC.frx":0E5E
      End
      Begin wizButton.cmd cmdVDR 
         Height          =   435
         Left            =   150
         TabIndex        =   2
         ToolTipText     =   "Vehicle Delivery Report"
         Top             =   661
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
         MICON           =   "Print_HGC.frx":0E7A
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
         MICON           =   "Print_HGC.frx":0E96
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
         MICON           =   "Print_HGC.frx":0EB2
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
         MICON           =   "Print_HGC.frx":0ECE
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
         MICON           =   "Print_HGC.frx":0EEA
      End
      Begin wizButton.cmd cmdClearance 
         Height          =   435
         Left            =   4560
         TabIndex        =   8
         ToolTipText     =   "PNP Motor Vehicle Clearance Application"
         Top             =   210
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
         MICON           =   "Print_HGC.frx":0F06
      End
      Begin wizButton.cmd cmdDR 
         Height          =   435
         Left            =   4560
         TabIndex        =   9
         Top             =   690
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
         MICON           =   "Print_HGC.frx":0F22
      End
      Begin wizButton.cmd cmd4 
         Height          =   435
         Left            =   150
         TabIndex        =   12
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
         MICON           =   "Print_HGC.frx":0F3E
      End
      Begin wizButton.cmd cmd6 
         Height          =   435
         Left            =   150
         TabIndex        =   13
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
         MICON           =   "Print_HGC.frx":0F5A
      End
      Begin wizButton.cmd cmdTransaction 
         Height          =   435
         Left            =   150
         TabIndex        =   14
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
         MICON           =   "Print_HGC.frx":0F76
      End
      Begin wizButton.cmd cmdjob 
         Height          =   435
         Left            =   150
         TabIndex        =   15
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
         MICON           =   "Print_HGC.frx":0F92
      End
   End
   Begin VB.PictureBox picCreditMemo 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   6165
      Left            =   0
      ScaleHeight     =   6165
      ScaleWidth      =   4410
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   4410
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   720
         ScaleHeight     =   1545
         ScaleWidth      =   2985
         TabIndex        =   16
         Top             =   1560
         Width           =   3015
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
            Height          =   345
            Left            =   360
            MaxLength       =   6
            TabIndex        =   20
            ToolTipText     =   "Input Credit Memo Serial Number"
            Top             =   360
            Width           =   2415
         End
         Begin wizButton.cmd cmdPrintCreditMemo 
            Height          =   435
            Left            =   480
            TabIndex        =   18
            Top             =   960
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
            MICON           =   "Print_HGC.frx":0FAE
         End
         Begin wizButton.cmd cmd5 
            Height          =   435
            Left            =   1560
            TabIndex        =   19
            Top             =   960
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
            MICON           =   "Print_HGC.frx":0FCA
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Memo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   840
            TabIndex        =   21
            Top             =   0
            Width           =   1275
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   3255
            _Version        =   655364
            _ExtentX        =   5741
            _ExtentY        =   450
            _StockProps     =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            VisualTheme     =   3
         End
      End
   End
   Begin VB.PictureBox picDebitMemo 
      BorderStyle     =   0  'None
      Height          =   5325
      Left            =   0
      ScaleHeight     =   5325
      ScaleWidth      =   6360
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   6360
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   720
         ScaleHeight     =   1425
         ScaleWidth      =   2865
         TabIndex        =   22
         Top             =   1560
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
            Height          =   315
            Left            =   240
            MaxLength       =   6
            TabIndex        =   27
            ToolTipText     =   "Input Debit Memo Serial Number"
            Top             =   360
            Width           =   2445
         End
         Begin wizButton.cmd cmdPrintDebitmemo 
            Height          =   435
            Left            =   480
            TabIndex        =   25
            Top             =   840
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
            MICON           =   "Print_HGC.frx":0FE6
         End
         Begin wizButton.cmd cmd7 
            Height          =   435
            Left            =   1590
            TabIndex        =   26
            Top             =   840
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
            MICON           =   "Print_HGC.frx":1002
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Debit Memo"
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
            Left            =   960
            TabIndex        =   24
            Top             =   0
            Width           =   1275
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   3255
            _Version        =   655364
            _ExtentX        =   5741
            _ExtentY        =   450
            _StockProps     =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            VisualTheme     =   3
         End
      End
   End
End
Attribute VB_Name = "frmSMIS_Report_Print_HGC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCustomer                                                        As ADODB.Recordset
Dim rsPurchAgree                                                      As ADODB.Recordset
Public GM                                                             As String
Public IGNKEYNO                                                       As String
Public VI_NO                                                          As String

Sub PRINTWARRANTYEXCEL(xYear)
    If Len(Dir(App.Path & "\warranty.xls")) = 0 Then
        MsgBox "Could Not located Warranty.Xls File", vbInformation
        Exit Sub
    End If
    On Error GoTo Errorcode
    Dim xlApp                                                         As Excel.Application
    Dim xlBook                                                        As Excel.Workbook
    Dim xlSheet                                                       As Excel.Worksheet
    Set xlApp = New Excel.Application

    Set xlBook = xlApp.Workbooks.Open(App.Path & "\warranty.xls")
    Set xlSheet = xlBook.Worksheets(1)
    Dim rsModel                                                       As ADODB.Recordset
    Dim vmodel                                                        As String
    Dim i                                                             As Integer
    Dim j                                                             As Integer
    Dim rsCountProspect                                               As ADODB.Recordset
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
Errorcode:
    MsgBox Err.Description
    Err.Clear
End Sub

Sub rsRefresh()
    Set rsPurchAgree = New ADODB.Recordset
    rsPurchAgree.Open "select * from SMIS_SALESORDER where code = '" & CUSCODE & "' AND IGNKEY_NO ='" & IGNKEYNO & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub cmd1_Click()
    Screen.MousePointer = 11
    rptPrint.Reset
    LoadSignatories ("GATE PASS")
    rptPrint.Formulas(0) = "ApprovedBy='" & Null2String(ApprovedBy) & "'"
    rptPrint.Formulas(1) = "ReleasedBy='" & Null2String(SalesDispatcher) & "'"
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "GatePass.rpt", "{SMIS_SalesOrder.VI_NO}='" & VI_NO & "'", DMIS_REPORT_Connection, 1
    '**************************
    NEW_LogAudit "V", "VEHICLE INVOICING", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "GatePass: INVOICE NO:" & VI_NO, "", ""
    '**************************
    Screen.MousePointer = 0
End Sub

Private Sub cmd4_Click()
    PRINTWARRANTYEXCEL IGNKEYNO
    '**************************
    NEW_LogAudit "V", "WARRANTY", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "Vehicle Invoicing: IGNEKY NO:" & IGNKEYNO, "", ""
    '**************************
End Sub

Private Sub cmd6_Click()
    frmSMIS_Files_Signatories.Show 1
End Sub

Private Sub cmdClearance_Click()
    rptPrint.Reset
    Screen.MousePointer = 11
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "clearance.rpt", "{customer.cuscde} = '" & CUSCODE & "' AND {PurchAgree.ProdNo} = '" & PRODUCTNO & "'", DMIS_REPORT_Connection, 1
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
    Set RSCREDIT = gconDMIS.Execute("select creditmemo from smis_salesorder where VI_NO='" & VI_NO & "'")
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
    Dim RSCREDIT                                                      As ADODB.Recordset
    Set RSCREDIT = gconDMIS.Execute("SELECT DEBITMEMO FROM SMIS_SALESORDER WHERE VI_NO='" & VI_NO & "'")
    If IsNull(RSCREDIT("DEBITMEMO").Value) = True Then
        txtDebitMemo = (GenerateCode("SMIS_SALESORDER", "DEBITMEMO", "000000"))
    Else
        txtDebitMemo = RSCREDIT("DEBITMEMO").Value
    End If
    Set RSCREDIT = Nothing
    picDebitMemo.Visible = True
    picMain.Visible = False
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

    rptPrint.Reset
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        Screen.MousePointer = 11
        LoadSignatories ("DELIVERY REPORT")
        rptPrint.Formulas(1) = "Company_Name = '" & COMPANY_NAME & "'"
        rptPrint.Formulas(2) = "Company_Address = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptPrint, SMIS_REPORT_PATH & "jobrequest.rpt", "{purchagree.VI_NO} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
        '**************************
        NEW_LogAudit "V", "VEHICLE INVOICING", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "Job Request Form: INVOICE NO:" & VI_NO, "", ""
        '**************************
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record Found!"
    End If
End Sub

Private Sub cmdPrintCreditMemo_Click()
    Screen.MousePointer = 11
    ''''''
    Dim lng                                                           As Integer
    lng = gconDMIS.Execute("select Count(*) from SMIS_SALESORDER WHERE CREDITMEMO=" & N2Str2Null(txtCreditMemo)).Fields(0).Value
    If lng >= 1 And UCase(Null2String(rsPurchAgree!CREDITMEMO)) <> UCase(txtCreditMemo) And Null2String(rsPurchAgree!CREDITMEMO) <> "" Then
        MessagePop RecSaveWarning, "Duplicate Record", "Credit Memo Number Already Exist"
        Screen.MousePointer = 0
        Exit Sub
    End If
    gconDMIS.Execute ("UPDATE SMIS_SALESORDER SET CREDITMEMO=" & N2Str2Null(txtCreditMemo) & " WHERE VI_NO='" & VI_NO & "'")
    rsRefresh
    rptPrint.WindowShowPrintBtn = True
    'PrintSQLReport rptPrint, SMIS_REPORT_PATH & "CREDITMEMO.rpt", "{MRR.ignkey} = '" & IGNKEYNO & "'", DMIS_REPORT_Connection, 1

    'Updated by: JUN 07/21/2008
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "CREDITMEMO.rpt", "{Purchagree.ignkey_no} = '" & IGNKEYNO & "'", DMIS_REPORT_Connection, 1

    '**************************
    NEW_LogAudit "V", "VEHICLE INVOICING", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "Credit Memo: INVOICE NO:" & VI_NO, "", ""
    '**************************
    Screen.MousePointer = 0
End Sub

Private Sub cmdPrintDebitmemo_Click()

    Screen.MousePointer = 11
    ''''''
    Dim lng                                                           As Integer
    lng = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_SALESORDER WHERE DEBITMEMO=" & N2Str2Null(txtDebitMemo)).Fields(0).Value
    If lng >= 1 And UCase(Null2String(rsPurchAgree!DEBITMEMO)) <> UCase(txtDebitMemo) And Null2String(rsPurchAgree!DEBITMEMO) <> "" Then
        MessagePop RecSaveWarning, "DUPLICATE RECORD", "DEBIT MEMO NUMBER ALREADY EXIST"
        Screen.MousePointer = 0
        Exit Sub
    End If
    gconDMIS.Execute ("UPDATE SMIS_SALESORDER SET DEBITMEMO=" & N2Str2Null(txtDebitMemo) & " WHERE VI_NO='" & VI_NO & "'")
    rsRefresh
    LoadSignatories ("DEBIT MEMO")
    rptPrint.Formulas(0) = "PreparedBy='" & Null2String(PreparedBy) & "'"
    rptPrint.Formulas(1) = "CheckedBy='" & Null2String(CheckedBy) & "'"
    rptPrint.Formulas(2) = "ApprovedBy='" & Null2String(ApprovedBy) & "'"
    rptPrint.Formulas(3) = "receivedby='" & Null2String(FinancingManager) & "'"
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "DEBITMEMO.RPT", "{PurchAgree.VI_NO} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
    '**************************
    NEW_LogAudit "V", "VEHICLE INVOICING", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "Debit Memo: INVOICE NO:" & VI_NO, "", ""
    '**************************
    Screen.MousePointer = 0
End Sub

Private Sub cmdReleaseOrder_Click()
    rptPrint.Reset

    Screen.MousePointer = 11
    rptPrint.Reset
    rptPrint.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptPrint.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "ReleaseOrder.rpt", "{purchagree.VI_NO} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
    '**************************
    NEW_LogAudit "V", "RELEASED ORDER", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "Released Order: INVOICE NO:" & VI_NO, "", ""
    '**************************
    Screen.MousePointer = 0
End Sub

Private Sub cmdTransaction_Click()
    rptPrint.Reset
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        Screen.MousePointer = 11
        LoadSignatories ("DELIVERY REPORT")
        rptPrint.Formulas(1) = "Company_Name = '" & COMPANY_NAME & "'"
        rptPrint.Formulas(2) = "Company_Address = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptPrint, SMIS_REPORT_PATH & "Transactionslip.rpt", "{purchagree.VI_NO} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1

        '**************************
        NEW_LogAudit "V", "VEHICLE INVOICING", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "Transaction Slip: INVOICE NO:" & VI_NO, "", ""
        '**************************

        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record Found!"
    End If
End Sub

Private Sub cmdVDR_Click()
    rptPrint.Reset
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        Screen.MousePointer = 11
        LoadSignatories ("DELIVERY REPORT")
        PrintSQLReport rptPrint, SMIS_REPORT_PATH & "vdr.rpt", "{purchagree.VI_NO} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
        '**************************
        NEW_LogAudit "V", "VEHICLE INVOICING", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "Vehicile Invoicing: VDR NO:" & Null2String(frmSMIS_Trans_VehicleInvoice.txtRelease_VDR.Text), "", ""
        '**************************
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record Found!"
    End If
End Sub

Private Sub cmdVI_Click()
    On Error GoTo Errorcode:
    rptPrint.Reset

    LoadSignatories ("SALES INVOICE")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then

        If Null2String(rsPurchAgree!TERM) = "COD" Then
            Screen.MousePointer = 11
            rptPrint.Formulas(0) = "PreparedBy='" & Null2String(PreparedBy) & "'"
            rptPrint.Formulas(1) = "CheckedBy='" & Null2String(CheckedBy) & "'"
            rptPrint.Formulas(2) = "ApprovedBy='" & Null2String(ApprovedBy) & "'"
            rptPrint.Formulas(3) = "GM='" & Null2String(GeneralManager) & "'"
            PrintSQLReport rptPrint, SMIS_REPORT_PATH & "vi.rpt", "{purchagree.VI_NO} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        ElseIf Null2String(rsPurchAgree!TERM) = "CPO" Then
            Screen.MousePointer = 11
            rptPrint.Formulas(0) = "PreparedBy='" & Null2String(PreparedBy) & "'"
            rptPrint.Formulas(1) = "CheckedBy='" & Null2String(CheckedBy) & "'"
            rptPrint.Formulas(2) = "ApprovedBy='" & Null2String(ApprovedBy) & "'"
            rptPrint.Formulas(3) = "GM='" & Null2String(GeneralManager) & "'"
            PrintSQLReport rptPrint, SMIS_REPORT_PATH & "vi_compo.rpt", "{purchagree.VI_NO} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
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
            rptPrint.Formulas(0) = "PreparedBy='" & Null2String(PreparedBy) & "'"
            rptPrint.Formulas(1) = "CheckedBy='" & Null2String(CheckedBy) & "'"
            rptPrint.Formulas(2) = "ApprovedBy='" & Null2String(ApprovedBy) & "'"
            rptPrint.Formulas(3) = "GM='" & Null2String(GeneralManager) & "'"
            PrintSQLReport rptPrint, SMIS_REPORT_PATH & "VI_FIN.rpt", "{purchagree.VI_NO} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        End If
    End If

    '**************************
    NEW_LogAudit "V", "VEHICLE INVOICING", "", FindTransactionID(VI_NO, "VI_NO", "SMIS_PURCHAGREE"), "", "INVOICE NO:" & VI_NO, "", ""
    '**************************
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    rsRefresh
    Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "select * from ALL_CUSTMASTER_SMIS where code = '" & CUSCODE & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
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

Private Sub txtCreditMemo_LostFocus()
    txtCreditMemo = Format(txtCreditMemo, "000000")
End Sub


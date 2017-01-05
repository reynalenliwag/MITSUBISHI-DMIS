VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmSMIS_Report_Print_HMH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing Process..."
   ClientHeight    =   4665
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
   Icon            =   "Print_HSR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   4290
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   4605
      Left            =   -45
      ScaleHeight     =   4605
      ScaleWidth      =   4530
      TabIndex        =   0
      Top             =   -15
      Width           =   4530
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
         MICON           =   "Print_HSR.frx":0E42
      End
      Begin wizButton.cmd cmdDebitMemo 
         Height          =   435
         Left            =   180
         TabIndex        =   5
         ToolTipText     =   "Debit Memo"
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
         MICON           =   "Print_HSR.frx":0E5E
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
         MICON           =   "Print_HSR.frx":0E7A
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
         MICON           =   "Print_HSR.frx":0E96
      End
      Begin wizButton.cmd cmdExit 
         Height          =   435
         Left            =   150
         TabIndex        =   6
         ToolTipText     =   "Exit"
         Top             =   4110
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
         MICON           =   "Print_HSR.frx":0EB2
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
         MICON           =   "Print_HSR.frx":0ECE
      End
      Begin wizButton.cmd cmdReleaseOrder 
         Height          =   435
         Left            =   150
         TabIndex        =   7
         ToolTipText     =   "Release Order"
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
         MICON           =   "Print_HSR.frx":0EEA
      End
      Begin wizButton.cmd cmd4 
         Height          =   435
         Left            =   150
         TabIndex        =   10
         ToolTipText     =   "Warranty Card Printing"
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
         MICON           =   "Print_HSR.frx":0F06
      End
      Begin wizButton.cmd cmd6 
         Height          =   435
         Left            =   150
         TabIndex        =   11
         ToolTipText     =   "Signatories Master File"
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
         MICON           =   "Print_HSR.frx":0F22
      End
   End
   Begin VB.PictureBox picDebitMemo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   4635
      Left            =   0
      ScaleHeight     =   4605
      ScaleWidth      =   4260
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   4290
      Begin VB.TextBox txtDebitMemo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   270
         MaxLength       =   6
         TabIndex        =   14
         ToolTipText     =   "Input Debit Memo Serial Number"
         Top             =   450
         Width           =   3735
      End
      Begin wizButton.cmd cmdPrintDebitmemo 
         Height          =   435
         Left            =   1920
         TabIndex        =   15
         Top             =   990
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
         MICON           =   "Print_HSR.frx":0F3E
      End
      Begin wizButton.cmd cmd7 
         Height          =   435
         Left            =   2970
         TabIndex        =   16
         Top             =   990
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
         MICON           =   "Print_HSR.frx":0F5A
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   345
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   4305
         _Version        =   655364
         _ExtentX        =   7594
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "Debit Memo"
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
   Begin VB.PictureBox picCreditMemo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   4665
      Left            =   0
      ScaleHeight     =   4635
      ScaleWidth      =   4260
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   4290
      Begin VB.TextBox txtCreditMemo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   300
         MaxLength       =   6
         TabIndex        =   17
         ToolTipText     =   "Input Credit Memo Serial Number"
         Top             =   510
         Width           =   3765
      End
      Begin wizButton.cmd cmdPrintCreditMemo 
         Height          =   435
         Left            =   1920
         TabIndex        =   18
         Top             =   1020
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
         MICON           =   "Print_HSR.frx":0F76
      End
      Begin wizButton.cmd cmd5 
         Height          =   435
         Left            =   3000
         TabIndex        =   19
         Top             =   1020
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
         MICON           =   "Print_HSR.frx":0F92
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   4455
         _Version        =   655364
         _ExtentX        =   7858
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Credit Memo"
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
Attribute VB_Name = "frmSMIS_Report_Print_HMH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCustomer                          As ADODB.Recordset
Dim rsPurchAgree                        As ADODB.Recordset
Public GM                               As String
Public IGNKEYNO                         As String
Public VI_NO                            As String

Private Sub cmd1_Click()
    Screen.MousePointer = 11
    rptPrint.Reset
    LoadSignatories ("GATE PASS")
    rptPrint.Formulas(0) = "ApprovedBy='" & Null2String(ApprovedBy) & "'"
    rptPrint.Formulas(1) = "ReleasedBy='" & Null2String(SalesDispatcher) & "'"
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "GatePass.rpt", "{SMIS_SalesOrder.VI_NO}='" & VI_NO & "'", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmd2_Click()

End Sub


Private Sub cmd4_Click()
    PRINTWARRANTYEXCEL IGNKEYNO
End Sub

Private Sub cmd5_Click()
    picMain.Visible = True: picCreditMemo.Visible = False: picCreditMemo.ZOrder 1
End Sub

Private Sub cmd6_Click()
    frmSMIS_Files_Signatories.Show 1
End Sub

Private Sub cmd7_Click()
    picMain.Visible = True: picDebitMemo.Visible = False: picDebitMemo.ZOrder 1
End Sub

Private Sub cmdCreditMemo_Click()
    Dim temprs                          As ADODB.Recordset
    rptPrint.Reset
    Set temprs = gconDMIS.Execute("select Count(*) from SMIS_MrrInv_Detail where  IgnKeyNo='" & IGNKEYNO & "'")
    If temprs(0).Value = 0 And rsPurchAgree!DISCOUNT = 0 Then
        MsgBox " There are No Record For This Transaction", vbInformation
        Exit Sub
    End If
    Dim RSCREDIT                        As ADODB.Recordset
    Set RSCREDIT = gconDMIS.Execute("select creditmemo from smis_salesorder where ignkey_no='" & IGNKEYNO & "'")
    If IsNull(RSCREDIT("CREDITMEMO").Value) = True Then
        txtCreditMemo = (GenerateCode("SMIS_SALESORDER", "CREDITMEMO", "000000"))
    Else
        txtCreditMemo = RSCREDIT("CREDITMEMO").Value
    End If
    picCreditMemo.Visible = True: picCreditMemo.ZOrder 0
    
End Sub

 
Private Sub cmdDebitMemo_Click()
    rptPrint.Reset
    Dim RSCREDIT                        As ADODB.Recordset
    Set RSCREDIT = gconDMIS.Execute("SELECT DEBITMEMO FROM SMIS_SALESORDER WHERE IGNKEY_NO='" & IGNKEYNO & "'")
    If IsNull(RSCREDIT("DEBITMEMO").Value) = True Then
        txtDebitMemo = (GenerateCode("SMIS_SALESORDER", "DEBITMEMO", "000000"))
    Else
        txtDebitMemo = RSCREDIT("DEBITMEMO").Value
    End If
    Set RSCREDIT = Nothing
    picDebitMemo.Visible = True
    picMain.Visible = False
End Sub

 

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrintCreditMemo_Click()
    Screen.MousePointer = 11
    Dim lng                             As Integer
    lng = gconDMIS.Execute("select Count(*) from SMIS_SALESORDER WHERE CREDITMEMO=" & N2Str2Null(txtCreditMemo)).Fields(0).Value
    If lng >= 1 And UCase(Null2String(rsPurchAgree!CREDITMEMO)) <> UCase(txtCreditMemo) And Null2String(rsPurchAgree!CREDITMEMO) <> "" Then
        MessagePop RecSaveWarning, "Duplicate Record", "Credit Memo Number Already Exist"
        Screen.MousePointer = 0
        Exit Sub
    End If
    gconDMIS.Execute ("UPDATE SMIS_SALESORDER SET CREDITMEMO=" & N2Str2Null(txtCreditMemo) & " WHERE IGNKEY_NO='" & IGNKEYNO & "'")
    rsRefresh
    rptPrint.WindowShowPrintBtn = True
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "CREDITMEMO.rpt", "{MRR.ignkey} = '" & IGNKEYNO & "'", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    
End Sub

Private Sub cmdPrintDebitmemo_Click()
    Screen.MousePointer = 11
    Dim lng                             As Integer
    lng = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_SALESORDER WHERE DEBITMEMO=" & N2Str2Null(txtDebitMemo)).Fields(0).Value
    If lng >= 1 And UCase(Null2String(rsPurchAgree!DEBITMEMO)) <> UCase(txtDebitMemo) And (Null2String(rsPurchAgree!DEBITMEMO) <> "") Then
        MessagePop RecSaveWarning, "DUPLICATE RECORD", "DEBIT MEMO NUMBER ALREADY EXIST"
        Screen.MousePointer = 0
        Exit Sub
    End If
    gconDMIS.Execute ("UPDATE SMIS_SALESORDER SET DEBITMEMO=" & N2Str2Null(txtDebitMemo) & " WHERE IGNKEY_NO='" & IGNKEYNO & "'")
    rsRefresh
    LoadSignatories ("DEBIT MEMO")
    rptPrint.Formulas(0) = "PreparedBy='" & Null2String(PreparedBy) & "'"
    rptPrint.Formulas(1) = "CheckedBy='" & Null2String(CheckedBy) & "'"
    rptPrint.Formulas(2) = "ApprovedBy='" & Null2String(ApprovedBy) & "'"
    rptPrint.Formulas(3) = "receivedby='" & Null2String(FinancingManager) & "'"



    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "DEBITMEMO.RPT", "{PurchAgree.IGNKEY_NO} = '" & IGNKEYNO & "'", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdReleaseOrder_Click()
    rptPrint.Reset
    Screen.MousePointer = 11
    rptPrint.Reset
    rptPrint.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptPrint.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptPrint, SMIS_REPORT_PATH & "ReleaseOrder.rpt", "{purchagree.ignKey_no} = '" & IGNKEYNO & "'", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdVDR_Click()
    rptPrint.Reset
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        Screen.MousePointer = 11
        LoadSignatories ("DELIVERY REPORT")
        rptPrint.Formulas(0) = "GM='" & Null2String(GeneralManager) & "'"
        PrintSQLReport rptPrint, SMIS_REPORT_PATH & "vdr.rpt", "{purchagree.vi_no} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record Found!"
    End If
End Sub

Private Sub cmdVI_Click()
    On Error GoTo ErrorCode:
    rptPrint.Reset
    LoadSignatories ("SALES INVOICE")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        If Null2String(rsPurchAgree!TERM) = "F" Or Null2String(rsPurchAgree!TERM) = "BPO" Then
            Screen.MousePointer = 11
            rptPrint.Formulas(0) = "GM='" & Null2String(GeneralManager) & "'"
            rptPrint.Formulas(1) = "PREPAREDBY='" & Null2String(PreparedBy) & "'"
            rptPrint.Formulas(2) = "CHECKEDBY='" & Null2String(CheckedBy) & "'"
            rptPrint.Formulas(3) = "APPROVEDBY='" & Null2String(ApprovedBy) & "'"
            PrintSQLReport rptPrint, SMIS_REPORT_PATH & "VI_FIN.rpt", "{purchagree.vi_no} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0

        Else
            Screen.MousePointer = 11
            rptPrint.Formulas(0) = "GM='" & Null2String(GeneralManager) & "'"
            rptPrint.Formulas(1) = "PREPAREDBY='" & Null2String(PreparedBy) & "'"
            rptPrint.Formulas(2) = "CHECKEDBY='" & Null2String(CheckedBy) & "'"
            rptPrint.Formulas(3) = "APPROVEDBY='" & Null2String(ApprovedBy) & "'"
            PrintSQLReport rptPrint, SMIS_REPORT_PATH & "vi.rpt", "{purchagree.vi_no} = '" & VI_NO & "'", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        End If
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
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
    VI_NO = ""
    PRODUCTNO = ""
    IGNKEYNO = ""
End Sub

Sub PRINTWARRANTYEXCEL(xYear)
    On Error GoTo ErrorCode
    Dim xlApp                           As Excel.Application
    Dim xlBook                          As Excel.Workbook
    Dim xlSheet                         As Excel.Worksheet
    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(App.Path & "\SMIS_EXCEL\WARRANTY.XLS")
    Set xlSheet = xlBook.Worksheets(1)
    Dim rsModel                         As ADODB.Recordset
    Dim vmodel                          As String
    Dim i                               As Integer
    Dim j                               As Integer
    Dim rsCountProspect                 As ADODB.Recordset
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
        
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    
    End If
    Exit Sub
ErrorCode:
    MsgBox Err.Description
    Err.Clear
End Sub

Sub rsRefresh()
    Set rsPurchAgree = New ADODB.Recordset
    rsPurchAgree.Open "select * from SMIS_SALESORDER where VI_NO='" & VI_NO & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub txtCreditMemo_LostFocus()
    txtCreditMemo = Format(txtCreditMemo, "000000")
End Sub


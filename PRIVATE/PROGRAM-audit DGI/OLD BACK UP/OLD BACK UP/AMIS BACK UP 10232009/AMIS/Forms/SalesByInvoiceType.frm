VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAMISSalesByInvoiceType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Income Statement"
   ClientHeight    =   1515
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   4830
   ForeColor       =   &H00FFFFFF&
   Icon            =   "SalesByInvoiceType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4830
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2445
      MouseIcon       =   "SalesByInvoiceType.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "SalesByInvoiceType.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   630
      Width           =   885
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1575
      MouseIcon       =   "SalesByInvoiceType.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "SalesByInvoiceType.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Report"
      Top             =   630
      Width           =   885
   End
   Begin Crystal.CrystalReport rptAMISIncomeStatement 
      Left            =   900
      Top             =   1020
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Income Statements"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   405
      Left            =   780
      TabIndex        =   1
      Top             =   90
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
      Format          =   51773441
      CurrentDate     =   38216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   405
      Left            =   3030
      TabIndex        =   3
      Top             =   90
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
      Format          =   51773441
      CurrentDate     =   38216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "From :"
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
      TabIndex        =   0
      Top             =   150
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
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
      Left            =   2550
      TabIndex        =   2
      Top             =   150
      Width           =   435
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   6
      Top             =   2970
      Width           =   495
   End
End
Attribute VB_Name = "frmAMISSalesByInvoiceType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                                      As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200714:10
Private Sub cmdPrint_Click()


    Dim Prev_dtpFrom, Prev_dtpTo                                      As String
    On Error GoTo Errorcode:

    If dtpFrom > dtpTo Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If
    If Month(dtpFrom) = 1 Then
        Prev_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
    Else
        Prev_dtpFrom = CDate(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom))
    End If
    If Month(dtpTo) = 1 Then
        Prev_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
    Else
        Prev_dtpTo = lastDay(Format(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom), "short date"))
    End If
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("select * from AMIS_Journal_HD where Jtype = 'SJ' and (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        Dim rsProfile                                                 As ADODB.Recordset
        rptAMISIncomeStatement.Reset
        rptAMISIncomeStatement.WindowShowPrintSetupBtn = True
        Set rsProfile = New ADODB.Recordset
        Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
        If Not (rsProfile.EOF And rsProfile.BOF) Then
            rptAMISIncomeStatement.Formulas(30) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
            rptAMISIncomeStatement.Formulas(31) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
            If INVOICE_Type = "VEHICLE" Then
                rptAMISIncomeStatement.ReportTitle = "VEHICLE SALES INVOICE"
                rptAMISIncomeStatement.WindowTitle = "VEHICLE SALES INVOICE"
            End If
            If INVOICE_Type = "H_VEHICLE" Then
                rptAMISIncomeStatement.ReportTitle = "HYUNDAI VEHICLE SALES INVOICE"
                rptAMISIncomeStatement.WindowTitle = "HYUNDAI VEHICLE SALES INVOICE"
            End If
            If INVOICE_Type = "PARTS-CASH" Then
                rptAMISIncomeStatement.ReportTitle = "PARTS CASH INVOICE"
                rptAMISIncomeStatement.WindowTitle = "PARTS CASH INVOICE"
            End If
            If INVOICE_Type = "PARTS-CHARGE" Then
                rptAMISIncomeStatement.ReportTitle = "PARTS CHARGE INVOICE"
                rptAMISIncomeStatement.WindowTitle = "PARTS CHARGE INVOICE"
            End If
            If INVOICE_Type = "H_PARTS-CASH" Then
                rptAMISIncomeStatement.ReportTitle = "HYUNDAI PARTS CASH INVOICE"
                rptAMISIncomeStatement.WindowTitle = "HYUNDAI PARTS CASH INVOICE"
            End If
            If INVOICE_Type = "H_PARTS-CHARGE" Then
                rptAMISIncomeStatement.ReportTitle = "HYUNDAI PARTS CHARGE INVOICE"
                rptAMISIncomeStatement.WindowTitle = "HYUNDAI PARTS CHARGE INVOICE"
            End If
            If INVOICE_Type = "SERVICE-CASH" Then
                rptAMISIncomeStatement.ReportTitle = "SALES SERVICE INVOICE - CASH"
                rptAMISIncomeStatement.WindowTitle = "SALES SERVICE INVOICE - CASH"
            End If
            If INVOICE_Type = "SERVICE-CHARGE" Then
                rptAMISIncomeStatement.ReportTitle = "SALES SERVICE INVOICE - CHARGE"
                rptAMISIncomeStatement.WindowTitle = "SALES SERVICE INVOICE - CHARGE"
            End If
            If INVOICE_Type = "H_SERVICE-CASH" Then
                rptAMISIncomeStatement.ReportTitle = "HYUNDAI SALES SERVICE INVOICE - CASH"
                rptAMISIncomeStatement.WindowTitle = "HYUNDAI SALES SERVICE INVOICE - CASH"
            End If
            If INVOICE_Type = "H_SERVICE-CHARGE" Then
                rptAMISIncomeStatement.ReportTitle = "HYUNDAI SALES SERVICE INVOICE - CHARGE"
                rptAMISIncomeStatement.WindowTitle = "HYUNDAI SALES SERVICE INVOICE - CHARGE"
            End If
        End If
        If INVOICE_Type = "VEHICLE" Then
            'rptAMISIncomeStatement.Reset
            'rptAMISIncomeStatement.ReportTitle = "VEHICLE SALES INVOICE"
            'rptAMISIncomeStatement.WindowTitle = "VEHICLE SALES INVOICE"
            'rptAMISIncomeStatement.ReportFileName = AMIS_REPORT_PATH & "InvoicesReport\VehicleInvoices.rpt"
            'rptAMISIncomeStatement.SelectionFormula = "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")"
            'rptAMISIncomeStatement.WindowState = crptMaximized
            'rptAMISIncomeStatement.Connect = DataEnvironment1.Connection1
            'rptAMISIncomeStatement.Connect = DMIS_REPORT_Connection
            'rptAMISIncomeStatement.Action = 1
            PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "InvoicesReport\VehicleInvoices.rpt", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
            LogAudit "V", "INCOME STATEMENT - VEHICLE SALES INVOICE", INVOICE_Type & ": " & dtpFrom & "-" & dtpTo
        End If
        If INVOICE_Type = "H_VEHICLE" Then
            'rptAMISIncomeStatement.Reset
            'rptAMISIncomeStatement.ReportTitle = "VEHICLE SALES INVOICE"
            'rptAMISIncomeStatement.WindowTitle = "VEHICLE SALES INVOICE"
            'rptAMISIncomeStatement.ReportFileName = AMIS_REPORT_PATH & "InvoicesReport\VehicleInvoices.rpt"
            'rptAMISIncomeStatement.SelectionFormula = "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")"
            'rptAMISIncomeStatement.WindowState = crptMaximized
            'rptAMISIncomeStatement.Connect = DataEnvironment1.Connection1
            'rptAMISIncomeStatement.Connect = DMIS_REPORT_Connection
            'rptAMISIncomeStatement.Action = 1
            PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "InvoicesReport\H_VehicleInvoices.rpt", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
            LogAudit "V", "INCOME STATEMENT - HYUNDAI VEHICLE SALES INVOICE", INVOICE_Type & ": " & dtpFrom & "-" & dtpTo
        End If
        If INVOICE_Type = "PARTS-CASH" Then
            PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "InvoicesReport\PartsCashInvoice.rpt", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
            LogAudit "V", "INCOME STATEMENT - PARTS CASH INVOICE", " & dtpFrom & " - " & dtpTo"
        End If
        If INVOICE_Type = "PARTS-CHARGE" Then
            PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "InvoicesReport\PartsChargeInvoice.rpt", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
            LogAudit "V", "INCOME STATEMENT - PARTS CHARGE INVOICE", " & dtpFrom & " - " & dtpTo"
        End If
        If INVOICE_Type = "H_PARTS-CASH" Then
            PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "InvoicesReport\H_PartsCashInvoice.rpt", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
            LogAudit "V", "INCOME STATEMENT - HYUNDAI PARTS CASH INVOICE", " & dtpFrom & " - " & dtpTo"
        End If
        If INVOICE_Type = "H_PARTS-CHARGE" Then
            PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "InvoicesReport\H_PartsChargeInvoice.rpt", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
            LogAudit "V", "INCOME STATEMENT - HYUNDAI PARTS CHARGE INVOICE", " & dtpFrom & " - " & dtpTo"
        End If
        If INVOICE_Type = "SERVICE-CASH" Then
            PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "InvoicesReport\ServiceInvoiceCash.rpt", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
            LogAudit "V", "INCOME STATEMENT - SALES SERVICE INVOICE - CASH", " & dtpFrom & " - " & dtpTo"
        End If
        If INVOICE_Type = "SERVICE-CHARGE" Then
            PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "InvoicesReport\ServiceInvoiceCharge.rpt", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
            LogAudit "V", "INCOME STATEMENT - SALES SERVICE INVOICE - CHARGE", " & dtpFrom & " - " & dtpTo"
        End If
        If INVOICE_Type = "H_SERVICE-CASH" Then
            PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "InvoicesReport\H_ServiceInvoiceCash.rpt", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
            LogAudit "V", "INCOME STATEMENT - HYUNDAI SALES SERVICE INVOICE - CASH", " & dtpFrom & " - " & dtpTo"
        End If
        If INVOICE_Type = "H_SERVICE-CHARGE" Then
            PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "InvoicesReport\H_ServiceInvoiceCharge.rpt", "{Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
            LogAudit "V", "INCOME STATEMENT - HYUNDAI SALES SERVICE INVOICE - CHARGE", " & dtpFrom & " - " & dtpTo"
        End If
    Else
        ShowNoRecord
    End If
    Call NEW_LogAudit("V", "INVOICE REPORT", "", "", "", dtpFrom & " " & dtpTo, "", "")
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
             
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (INVOICE REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "INVOICE REPORT", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    If INVOICE_Type = "VEHICLE" Then
        Me.Caption = "VEHICLES SALES INVOICE"
    End If
    If INVOICE_Type = "H_VEHICLE" Then
        Me.Caption = "HYUNDAI VEHICLES SALES INVOICE"
    End If
    If INVOICE_Type = "PARTS-CASH" Then
        Me.Caption = "PARTS CASH INVOICE"
    End If
    If INVOICE_Type = "PARTS-CHARGE" Then
        Me.Caption = "PARTS CHARGE INVOICE"
    End If
    If INVOICE_Type = "H_PARTS-CASH" Then
        Me.Caption = "HYUNDAI PARTS CASH INVOICE"
    End If
    If INVOICE_Type = "H_PARTS-CHARGE" Then
        Me.Caption = "HYUNDAI PARTS CHARGE INVOICE"
    End If
    If INVOICE_Type = "SERVICE-CASH" Then
        Me.Caption = "SALES SERVICE INVOICE - CASH"
    End If
    If INVOICE_Type = "SERVICE-CHARGE" Then
        Me.Caption = "SALES SERVICE INVOICE - CHARGE"
    End If
    If INVOICE_Type = "H_SERVICE-CASH" Then
        Me.Caption = "HYUNDAI SALES SERVICE INVOICE - CASH"
    End If
    If INVOICE_Type = "H_SERVICE-CHARGE" Then
        Me.Caption = "HYUNDAI SALES SERVICE INVOICE - CHARGE"
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAMISRange = Nothing
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub


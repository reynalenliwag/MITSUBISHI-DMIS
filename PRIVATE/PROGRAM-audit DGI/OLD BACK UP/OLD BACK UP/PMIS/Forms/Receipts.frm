VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPMISReports_Receipts 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Receipts"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Receipts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2310
   ScaleWidth      =   4365
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
      Height          =   795
      Left            =   2040
      MouseIcon       =   "Receipts.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "Receipts.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   1380
      Width           =   675
   End
   Begin VB.ComboBox cboSupplier 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   60
      TabIndex        =   0
      Text            =   "cboSupplier"
      ToolTipText     =   "Select month from the list"
      Top             =   120
      Width           =   4245
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select month from the list"
      Top             =   540
      Width           =   1965
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Select year from the list"
      Top             =   930
      Width           =   1965
   End
   Begin Crystal.CrystalReport rptReceipts 
      Left            =   210
      Top             =   1350
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Monthly Receipts"
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
   Begin VB.TextBox txtSupCode 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2580
      Width           =   555
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
      Height          =   795
      Left            =   1380
      MouseIcon       =   "Receipts.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "Receipts.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   1380
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   540
      TabIndex        =   7
      Top             =   570
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   540
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   5
      Top             =   3060
      Width           =   495
   End
End
Attribute VB_Name = "frmPMISReports_Receipts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsREC_HIST                                                        As ADODB.Recordset
Dim rsSupplier                                                        As ADODB.Recordset

Private Sub cboSupplier_Click()
    Set rsSupplier = New ADODB.Recordset
    Set rsSupplier = gconDMIS.Execute("Select * from PMIS_vw_Supplier where SupName = '" & cboSupplier.Text & "'")
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        txtSupCode.Text = Null2String(rsSupplier!SupCode)
    End If
End Sub

Private Sub cboSupplier_GotFocus()
    VBComBoBoxDroppedDown cboSupplier
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:

    Set rsREC_HIST = New ADODB.Recordset
    rsREC_HIST.Open "select RRdate from PMIS_Rec_Hist where TYPE = 'P' AND month(RRDate) = " & What_month(cboMonth) & " AND year(RRDate) =" & cboYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsREC_HIST.EOF And Not rsREC_HIST.EOF Then
        Screen.MousePointer = 11
        If cboSupplier.Text = "ALL SUPPLIERS" Then
            rptReceipts.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptReceipts.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptReceipts, PMIS_REPORT_PATH & "receipts.rpt", "{RR_HD.TYPE} = 'P' AND month({RR_HD.RRDate}) = " & What_month(cboMonth.Text) & " AND year({RR_HD.RRDate}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
            PrintSQLReport rptReceipts, PMIS_REPORT_PATH & "receiptsum.rpt", "{RR_HD.TYPE} = 'P' AND month({RR_HD.RRDate}) = " & What_month(cboMonth.Text) & " AND year({RR_HD.RRDate}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        Else
            rptReceipts.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptReceipts.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptReceipts, PMIS_REPORT_PATH & "receipts.rpt", "{RR_HD.TYPE} = 'P' AND {RR_HD.RECVD_CODE} = '" & txtSupCode.Text & "' AND month({RR_HD.RRDate}) = " & What_month(cboMonth.Text) & " AND year({RR_HD.RRDate}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
            PrintSQLReport rptReceipts, PMIS_REPORT_PATH & "receiptsum.rpt", "{RR_HD.TYPE} = 'P' AND {RR_HD.RECVD_CODE} = '" & txtSupCode.Text & "' AND month({RR_HD.RRDate}) = " & What_month(cboMonth.Text) & " AND year({RR_HD.RRDate}) = " & cboYear.Text, DMIS_REPORT_Connection, 1
        End If
        NEW_LogAudit "V", "RECEIPT FOR THE MONTH", "", "", "", cboMonth & " - " & cboYear & " " & cboSupplier, "", ""
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record for the Month of " & cboMonth.Text & " Year " & cboYear.Text
    End If
    
    Exit Sub
    
ErrorCode:
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (RECEIPT FOR THE MONTH)"
            Call frmALL_AuditInquiry.DisplayHistory("", "RECEIPT FOR THE MONTH", "PRINTING")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonth
    FillcboYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    cboSupplier.Clear
    cboSupplier.AddItem "ALL SUPPLIERS"
    Set rsSupplier = New ADODB.Recordset
    Set rsSupplier = gconDMIS.Execute("Select * from PMIS_vw_Supplier Order By SupName asc")
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        rsSupplier.MoveFirst
        Do While Not rsSupplier.EOF
            cboSupplier.AddItem Null2String(rsSupplier!supname)
            rsSupplier.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISReports_Receipts = Nothing
    UnloadForm Me
End Sub


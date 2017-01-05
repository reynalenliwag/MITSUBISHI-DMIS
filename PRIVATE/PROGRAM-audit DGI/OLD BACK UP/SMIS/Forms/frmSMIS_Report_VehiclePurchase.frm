VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_VehiclePurchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle purchase Report"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   Icon            =   "frmSMIS_Report_VehiclePurchase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1830
   ScaleWidth      =   4830
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   465
      ItemData        =   "frmSMIS_Report_VehiclePurchase.frx":08CA
      Left            =   30
      List            =   "frmSMIS_Report_VehiclePurchase.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   90
      Width           =   2355
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   465
      Left            =   3390
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   90
      Width           =   1365
   End
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
      Left            =   2280
      MouseIcon       =   "frmSMIS_Report_VehiclePurchase.frx":08CE
      MousePointer    =   99  'Custom
      Picture         =   "frmSMIS_Report_VehiclePurchase.frx":0A20
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close Window"
      Top             =   870
      Width           =   885
   End
   Begin Crystal.CrystalReport rptReleased 
      Left            =   4320
      Top             =   1740
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "MMPC Purchases Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
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
      Left            =   1410
      MouseIcon       =   "frmSMIS_Report_VehiclePurchase.frx":0E6B
      MousePointer    =   99  'Custom
      Picture         =   "frmSMIS_Report_VehiclePurchase.frx":0FBD
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   870
      Width           =   885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   2490
      TabIndex        =   4
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "frmSMIS_Report_VehiclePurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMRRINV                                                          As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:
    Set rsMRRINV = New ADODB.Recordset
    rsMRRINV.Open "select * from SMIS_po WHERE STATUS='P' AND  year(DATEORDERED) = '" & cboYear.Text & "' and month(DATEORDERED) = " & What_month(cboMonth), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
        Screen.MousePointer = 11
        rptReleased.WindowShowGroupTree = False
        rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptReleased.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptReleased.WindowTitle = "Monthly Vehicle Purchases Report"
        
        PrintSQLReport rptReleased, SMIS_REPORT_PATH & "POLISTING.rpt", "{PO.STATUS}='P' and year({PO.DATEORDERED}) = " & cboYear.Text & " AND month({PO.DATEORDERED}) = " & What_month(cboMonth), DMIS_REPORT_Connection, 1
        'UPDATED BY: JUN
        'DATE UPDATED: 09032008 5:00
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "PURCHASE ORDER REPORT", "", "", "", cboMonth & " " & cboYear, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

        'LogAudit "V", "VEHICLE PURCHASE REPORT", "FOR THE MONTH OF " & cboMonth & " YEAR " & cboYear
        Screen.MousePointer = 0
    Else

        MsgSpeechBox "No Record for " & cboMonth.Text & " " & cboYear.Text
    End If
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (PURCHASE ORDER REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "PURCHASE ORDER REPORT", "PRINTING")
    End Select

End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonth
    FillCboMoreYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub

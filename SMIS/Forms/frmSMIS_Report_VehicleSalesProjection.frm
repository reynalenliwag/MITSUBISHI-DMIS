VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_VehicleSalesProjection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VEHICLE SALES PROJECTION"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmSMIS_Report_VehicleSalesProjection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4680
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
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   180
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
      Left            =   3300
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   180
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
      Left            =   2145
      MouseIcon       =   "frmSMIS_Report_VehicleSalesProjection.frx":3482
      MousePointer    =   99  'Custom
      Picture         =   "frmSMIS_Report_VehicleSalesProjection.frx":35D4
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close Window"
      Top             =   780
      Width           =   885
   End
   Begin Crystal.CrystalReport rptReleased 
      Left            =   3450
      Top             =   1140
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Units Released"
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
      Left            =   1275
      MouseIcon       =   "frmSMIS_Report_VehicleSalesProjection.frx":3A1F
      MousePointer    =   99  'Custom
      Picture         =   "frmSMIS_Report_VehicleSalesProjection.frx":3B71
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print Report"
      Top             =   780
      Width           =   885
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
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
      Left            =   2430
      TabIndex        =   4
      Top             =   210
      Width           =   825
   End
End
Attribute VB_Name = "frmSMIS_Report_VehicleSalesProjection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Updated by: JUN
'Date Updated: 10082008
Option Explicit
Dim rsPurchAgree                                                      As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:

    If Len(cboYear.Text) = 4 Or cboYear.Text <> "" Then
        Set rsPurchAgree = New ADODB.Recordset
        rsPurchAgree.Open "select * from SMIS_MrrInv WHERE year(datereleased) = " & cboYear.Text, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPurchAgree.EOF And Not rsPurchAgree.BOF Then
            Screen.MousePointer = 11
            rptReleased.WindowShowGroupTree = False
            rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptReleased.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptReleased.Formulas(2) = "curr_month = Cdate('" & DateSerial(cboYear, What_month(cboMonth), 1) & "')"
            PrintSQLReport rptReleased, SMIS_REPORT_PATH & "VS\VehiclesSalesProjection.rpt", "", DMIS_REPORT_Connection, 1
            Call NEW_LogAudit("V", "SALES REPORT", "", "", "", cboMonth & " " & cboYear, "", "")
            Screen.MousePointer = 0
        Else
            MsgSpeechBox "No Record for " & cboMonth.Text & " " & cboYear.Text
        End If
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (VEHICLE SALES PROJECTION)"
            Call frmALL_AuditInquiry.DisplayHistory("", "VEHICLE SALES PROJECTION", "PRINTING")
            
    End Select

End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonth
    fillcbomoreyear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub



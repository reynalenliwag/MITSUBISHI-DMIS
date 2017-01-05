VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_UnitCommission 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Unit Commisson"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   ForeColor       =   &H00FCFCFC&
   Icon            =   "Report_UnitCommission.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1710
   ScaleWidth      =   5010
   Begin VB.CommandButton Command2 
      Caption         =   "Vehicles Group Listing"
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
      Left            =   90
      MouseIcon       =   "Report_UnitCommission.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "Report_UnitCommission.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Customer Directory by Customer Type"
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
      Left            =   60
      MouseIcon       =   "Report_UnitCommission.frx":08FB
      MousePointer    =   99  'Custom
      Picture         =   "Report_UnitCommission.frx":0A4D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1740
      Width           =   4575
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
      Left            =   2475
      MouseIcon       =   "Report_UnitCommission.frx":0EEC
      MousePointer    =   99  'Custom
      Picture         =   "Report_UnitCommission.frx":103E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   780
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
      Left            =   1605
      MouseIcon       =   "Report_UnitCommission.frx":1489
      MousePointer    =   99  'Custom
      Picture         =   "Report_UnitCommission.frx":15DB
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   780
      Width           =   885
   End
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
      ItemData        =   "Report_UnitCommission.frx":1A7A
      Left            =   90
      List            =   "Report_UnitCommission.frx":1A7C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2355
   End
   Begin Crystal.CrystalReport rpCom 
      Left            =   4050
      Top             =   1050
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "MMPC Monthly Inventory Control"
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
   Begin VB.TextBox txtYear 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   555
      Left            =   3450
      MaxLength       =   4
      TabIndex        =   2
      Text            =   "9999"
      Top             =   90
      Width           =   1005
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
      Left            =   2610
      TabIndex        =   1
      Top             =   150
      Width           =   825
   End
End
Attribute VB_Name = "frmSMIS_Report_UnitCommission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUnitCommission                                                  As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()


    Screen.MousePointer = 11

    On Error GoTo ErrorCode




    If Len(txtYear.Text) = 4 Or txtYear.Text <> "" Then
        Set rsUnitCommission = New ADODB.Recordset
        rsUnitCommission.Open "select * from SMIS_SalesOrder WHERE year(DateReleased) = '" & txtYear.Text & "' and month(DateReleased) = " & What_month(cboMonth), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsUnitCommission.EOF And Not rsUnitCommission.BOF Then
            Screen.MousePointer = 11
            rpCom.Reset
            rpCom.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rpCom.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rpCom.Formulas(2) = "forthe = '" & cboMonth.Text & " " & txtYear & "'"
            PrintSQLReport rpCom, SMIS_REPORT_PATH & "UNITCOMMISION.rpt", " Month({SO.DateReleased})=" & What_month(cboMonth.Text) & " AND Year({SO.DateReleased})=" & txtYear, DMIS_REPORT_Connection, 1
            'UPDATED BY: JUN
            'DATE UPDATED: 09032008 5:00
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
             Call NEW_LogAudit("V", "UNIT COMMISSION", "", "", "", cboMonth & " " & txtYear, "", "")
            'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------

            'LogAudit "V", "UNIT COMMISSION REPORT", "FOR THE MONTH OF " & cboMonth & " YEAR " & txtYear
        Else
            ShowNoRecord
        End If

        'End of update
    End If
    Screen.MousePointer = 0

    Exit Sub

ErrorCode:
    ShowVBError
    Screen.MousePointer = 0


End Sub

Private Sub Command2_Click()
    Screen.MousePointer = 11

    rpCom.WindowTitle = "Customer Directory By Customer Type"
    rpCom.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rpCom.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rpCom, SMIS_REPORT_PATH & "VehiclesGroupList.rpt", "", DMIS_REPORT_Connection, 1

    Screen.MousePointer = 0
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (UNIT COMMISSION)"
            Call frmALL_AuditInquiry.DisplayHistory("", "UNIT COMMISSION", "PRINTING")
            
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonth
    cboMonth.Text = The_month(Month(LOGDATE))
    txtYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub


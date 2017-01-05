VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSMIS_Report_Purchases2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receiving Report "
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   ForeColor       =   &H00FCFCFC&
   Icon            =   "ReportPurchases2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "ReportPurchases2.frx":0E42
   ScaleHeight     =   2745
   ScaleWidth      =   4725
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   4725
      TabIndex        =   7
      Top             =   840
      Width           =   4725
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   405
         Left            =   1470
         TabIndex        =   8
         Top             =   90
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56819713
         CurrentDate     =   38216
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   405
         Left            =   1470
         TabIndex        =   9
         Top             =   540
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56819713
         CurrentDate     =   38216
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   630
         Width           =   435
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   690
         TabIndex        =   10
         Top             =   90
         Width           =   675
      End
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Select Date Ranged"
      Height          =   375
      Index           =   1
      Left            =   30
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   390
      Width           =   4665
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Monthly Evaluation"
      Height          =   375
      Index           =   0
      Left            =   30
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   30
      Width           =   4665
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
      Left            =   2490
      MouseIcon       =   "ReportPurchases2.frx":1184
      MousePointer    =   99  'Custom
      Picture         =   "ReportPurchases2.frx":12D6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   1890
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
      Left            =   1620
      MouseIcon       =   "ReportPurchases2.frx":1721
      MousePointer    =   99  'Custom
      Picture         =   "ReportPurchases2.frx":1873
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   1890
      Width           =   885
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
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   1365
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
      ItemData        =   "ReportPurchases2.frx":1D12
      Left            =   60
      List            =   "ReportPurchases2.frx":1D14
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   2355
   End
   Begin Crystal.CrystalReport rptReleased 
      Left            =   3540
      Top             =   1080
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
      Left            =   2520
      TabIndex        =   2
      Top             =   990
      Width           =   825
   End
End
Attribute VB_Name = "frmSMIS_Report_Purchases2"
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


    If Opt(0).Value = True Then
            Set rsMRRINV = New ADODB.Recordset
            'rsMRRINV.Open "select * from SMIS_MrrInv WHERE year(datereleased) = '" & cboYear.Text & "' and month(datereceived) = " & What_month(cboMonth), gconDMIS, adOpenForwardOnly, adLockReadOnly
            rsMRRINV.Open "select * from SMIS_MrrInv WHERE year(datereceived) = '" & cboYear.Text & "' and month(datereceived) = " & What_month(cboMonth), gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
                Screen.MousePointer = 11
                rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptReleased.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptReleased.WindowTitle = "Monthly Vehicle Purchase Report"
                rptReleased.WindowShowGroupTree = False
                PrintSQLReport rptReleased, SMIS_REPORT_PATH & "received2.rpt", "year({VEHICLE.datereceived}) = " & cboYear.Text & " AND month({VEHICLE.datereceived}) = " & What_month(cboMonth), DMIS_REPORT_Connection, 1
                
                'NEW LOG AUDIT-----------------------------------------------------
                         Call NEW_LogAudit("V", "MONTHLY PURCHASE REPORT", "", "", "", "FROM " & " " & cboMonth & " " & cboYear, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
        
                Screen.MousePointer = 0
            Else
        
                MsgSpeechBox "No Record for " & cboMonth.Text & " " & cboYear.Text
            End If
      Else
            Dim CRYS_FILTER As String
            Dim SQL As String
            Set rsMRRINV = New ADODB.Recordset
            SQL = "SELECT * from SMIS_MrrInv"
            
            Set rsMRRINV = New ADODB.Recordset
            Set rsMRRINV = gconDMIS.Execute(SQL)
            
            If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
                Screen.MousePointer = 11
                rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptReleased.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptReleased.WindowTitle = "Monthly Vehicle Purchase Report"
                rptReleased.WindowShowGroupTree = False
                CRYS_FILTER = "{VEHICLE.datereceived} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {VEHICLE.datereceived} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")"
                PrintSQLReport rptReleased, SMIS_REPORT_PATH & "Customreports/received2.rpt", CRYS_FILTER, SMIS_REPORT_CONNECTION, 1
                Call NEW_LogAudit("V", "MONTHLY PURCHASE REPORT", "", "", "", "FROM " & " " & cboMonth & " " & cboYear, "", "")
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (MONTHLY PURCHASE REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "MONTHLY PURCHASE REPORT", "PRINTING")
            
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
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    Screen.MousePointer = 0
End Sub

Private Sub Opt_Click(Index As Integer)
If Opt(0).Value = True Then
    Picture1.Visible = False
    Opt(0).BackColor = &HFFFF80
    Opt(1).BackColor = &H8000000F
    Else
    Picture1.Visible = True
    Opt(1).BackColor = &HFFFF80
    Opt(0).BackColor = &H8000000F
    
End If
End Sub

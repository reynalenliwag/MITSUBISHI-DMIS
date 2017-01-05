VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_Monthly 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Customer / Customer Type "
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   ForeColor       =   &H00FCFCFC&
   Icon            =   "Report_CustomerCustomerType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
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
      MouseIcon       =   "Report_CustomerCustomerType.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "Report_CustomerCustomerType.frx":045C
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
      MouseIcon       =   "Report_CustomerCustomerType.frx":08FB
      MousePointer    =   99  'Custom
      Picture         =   "Report_CustomerCustomerType.frx":0A4D
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
      MouseIcon       =   "Report_CustomerCustomerType.frx":0EEC
      MousePointer    =   99  'Custom
      Picture         =   "Report_CustomerCustomerType.frx":103E
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
      MouseIcon       =   "Report_CustomerCustomerType.frx":1489
      MousePointer    =   99  'Custom
      Picture         =   "Report_CustomerCustomerType.frx":15DB
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
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2355
   End
   Begin Crystal.CrystalReport rptReleased 
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
Attribute VB_Name = "frmSMIS_Report_Monthly"
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

    Dim FILTER                                                        As String
    On Error GoTo ErrorCode:

    If Len(txtYear.Text) = 4 Or txtYear.Text <> "" Then
        Set rsMRRINV = New ADODB.Recordset
        rsMRRINV.Open "select * from SMIS_MrrInv WHERE year(datereceived) = '" & txtYear.Text & "' and month(datereceived) <= " & What_month(cboMonth), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
            Screen.MousePointer = 11
            FILTER = "(year({VEHICLE.datereceived}) = " & txtYear.Text & " AND month({VEHICLE.datereceived}) <= " & What_month(cboMonth) & " AND {VEHICLE.Released} = false)" & _
                   " OR " & _
                     "(year({VEHICLE.datereceived}) = " & txtYear.Text & " AND month({VEHICLE.datereceived}) <= " & What_month(cboMonth) & " AND {VEHICLE.Released} = true AND month({VEHICLE.datereleased}) > " & What_month(cboMonth) & ")"
            rptReleased.WindowTitle = "Vehicle Inventory Report"
            PrintSQLReport rptReleased, SMIS_REPORT_PATH & "invcontrol.rpt", FILTER, DMIS_REPORT_Connection, 1

            Screen.MousePointer = 0
        Else
            MsgSpeechBox "No Record for " & cboMonth.Text & " " & txtYear.Text
            Exit Sub
        End If
    End If





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command1_Click()
    Set rsMRRINV = New ADODB.Recordset
    rsMRRINV.Open "select * from ALL_CUSTOMER ", gconDMIS, adOpenForwardOnly, adLockReadOnly


    If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
        Screen.MousePointer = 11

        rptReleased.WindowTitle = "Customer Directory By Customer Type"
        rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptReleased.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptReleased, SMIS_REPORT_PATH & "CustomerListByCustType.rpt", "", DMIS_REPORT_Connection, 1

        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Record for " & cboMonth.Text & " " & txtYear.Text
        Exit Sub
    End If


End Sub

Private Sub Command2_Click()
    Screen.MousePointer = 11

    rptReleased.WindowTitle = "Customer Directory By Customer Type"
    rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptReleased.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptReleased, SMIS_REPORT_PATH & "listing/VehiclesGroupList.rpt", "", DMIS_REPORT_Connection, 1

    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    fillcbomonth cboMonth
    cboMonth.Text = The_month(Month(LOGDATE))
    txtYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub


VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSMIS_Report_YearlySalesAnalysis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yearly Sales Analysis"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3315
   ForeColor       =   &H00FCFCFC&
   Icon            =   "Report_YearlySalesAnalysis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3315
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
      Left            =   1710
      MouseIcon       =   "Report_YearlySalesAnalysis.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "Report_YearlySalesAnalysis.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Close Window"
      Top             =   660
      Width           =   885
   End
   Begin Crystal.CrystalReport rptReleased 
      Left            =   2310
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Yearly Ending Inventory"
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
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00973640&
      Height          =   405
      Left            =   1140
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "9999"
      Top             =   120
      Width           =   1065
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
      Left            =   840
      MouseIcon       =   "Report_YearlySalesAnalysis.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "Report_YearlySalesAnalysis.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print Report"
      Top             =   660
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   450
      TabIndex        =   1
      Top             =   180
      Width           =   510
   End
End
Attribute VB_Name = "frmSMIS_Report_YearlySalesAnalysis"
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
    On Error GoTo ErrorCode:

    If Len(txtYear.Text) = 4 Or txtYear.Text <> "" Then
        Set rsMRRINV = New ADODB.Recordset
        rsMRRINV.Open "select * from SMIS_MrrInv WHERE year(datereceived) = '" & txtYear.Text & "' AND released = 0", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
            Screen.MousePointer = 11
            rptReleased.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptReleased.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptReleased, SMIS_REPORT_PATH & "invcontrol2.rpt", "(year({VEHICLE.datereceived}) = " & txtYear.Text & " AND {VEHICLE.Released} = false)", DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        Else
            MsgSpeechBox "No Record for " & txtYear.Text
            Exit Sub
        End If
    End If

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtYear.Text = Year(LOGDATE)
    Screen.MousePointer = 0
End Sub


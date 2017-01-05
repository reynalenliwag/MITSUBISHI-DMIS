VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmPMISReports_SlowMoving 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MOVEMENT CATEGORY"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   ForeColor       =   &H00DEDFDE&
   Icon            =   "SlowMoving.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3810
   ScaleWidth      =   4695
   Begin VB.CheckBox chkDontIncludeLocal 
      BackColor       =   &H00DEDFDE&
      Caption         =   "Don't Include Non Mitsubishi Parts (Local Parts)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   390
      TabIndex        =   7
      Top             =   3810
      Width           =   4215
   End
   Begin VB.ComboBox cboDate_Gen 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select date form the list"
      Top             =   120
      Width           =   2205
   End
   Begin Crystal.CrystalReport rptRanks 
      Left            =   5670
      Top             =   6690
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin wizButton.cmd cmdD 
      Height          =   495
      Left            =   180
      TabIndex        =   1
      ToolTipText     =   "View detailed list of Rank D Parts"
      Top             =   540
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      TX              =   "RANK D  - SLOW MOVING PARTS      "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "SlowMoving.frx":0E42
   End
   Begin wizButton.cmd cmdE1 
      Height          =   495
      Left            =   180
      TabIndex        =   2
      ToolTipText     =   "View detailed list of Rank E1 Parts"
      Top             =   1080
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      TX              =   "RANK E1 - NON-MOVING FOR 1 YEAR  "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "SlowMoving.frx":115C
   End
   Begin wizButton.cmd cmdE2 
      Height          =   495
      Left            =   180
      TabIndex        =   3
      ToolTipText     =   "View detailed list of Rank E2 Parts"
      Top             =   1620
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      TX              =   "RANK E2 - NON-MOVING FOR 2 YEARS "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "SlowMoving.frx":1476
   End
   Begin wizButton.cmd cmdE3 
      Height          =   495
      Left            =   180
      TabIndex        =   4
      ToolTipText     =   "View detailed list of Rank E3 Parts"
      Top             =   2160
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      TX              =   "RANK E3 - NON-MOVING FOR 3 YEARS "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "SlowMoving.frx":1790
   End
   Begin wizButton.cmd cmdE4 
      Height          =   495
      Left            =   180
      TabIndex        =   5
      ToolTipText     =   "View detailed list of Rank E4 Parts"
      Top             =   2700
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      TX              =   "RANK E4 - NON-MOVING FOR 4 YEARS "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "SlowMoving.frx":1AAA
   End
   Begin wizButton.cmd cmdE5 
      Height          =   495
      Left            =   180
      TabIndex        =   6
      ToolTipText     =   "View detailed list of Rank E5 Parts"
      Top             =   3240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      TX              =   "RANK E5 - NON-MOVING FOR 5 YEARS "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "SlowMoving.frx":1DC4
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AS OF:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   750
      TabIndex        =   8
      Top             =   150
      Width           =   1035
   End
End
Attribute VB_Name = "frmPMISReports_SlowMoving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsRANKFLE                                          As ADODB.Recordset

Sub PrintInventoryRank(vINVCLASS As String, vSUBINVCLAS As String)
    If Function_Access(LOGID, "Acess_Print", "REPORTS INTERNAL SLOW MOVING PARTS FOR DISPOSAL") = False Then Exit Sub
    Dim ReportTitle                                    As String
    ReportTitle = "RANK " & vINVCLASS & vSUBINVCLAS & " SLOW MOVING PARTS WITH SELLING PRICE"
    vINVCLASS = "'" & vINVCLASS & "'"
    If vSUBINVCLAS = "" Then vSUBINVCLAS = "NULL" Else vSUBINVCLAS = "'" & vSUBINVCLAS & "'"
    Screen.MousePointer = 11
    rptRanks.ReportTitle = ReportTitle
    If chkDontIncludeLocal.Value = 1 Then
        If vSUBINVCLAS = "NULL" Then
            rptRanks.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptRanks.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "SlowMoving.rpt", "MID({rankfle.STOCKNO},2,3) <> 'NPN' AND {rankfle.invclass}=" & vINVCLASS & " and isnull({rankfle.subinvclas})= true AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
            rptRanks.ReportTitle = Replace(Replace("RANK " & vINVCLASS & vSUBINVCLAS & " SLOW MOVING ACCESSORIES WITH SELLING PRICE", "NULL", ""), "'", "")
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "SlowMoving_A.rpt", "MID({rankfle.STOCKNO},2,3) <> 'NPN' AND {rankfle.invclass}=" & vINVCLASS & " and isnull({rankfle.subinvclas})= true AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
            rptRanks.ReportTitle = Replace(Replace("RANK " & vINVCLASS & vSUBINVCLAS & "  SLOW MOVING MATERIALS WITH SELLING PRICE", "NULL", ""), "'", "")
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "SlowMoving_M.rpt", "MID({rankfle.STOCKNO},2,3) <> 'NPN' AND {rankfle.invclass}=" & vINVCLASS & " and isnull({rankfle.subinvclas})= true AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
            rptRanks.ReportTitle = Replace(Replace("RANK " & vINVCLASS & vSUBINVCLAS & "  SLOW MOVING PARTS,ACCESSORIES,MATERIALS WITH SELLING PRICE", "NULL", ""), "'", "")
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "SlowMoving_All.rpt", "MID({rankfle.STOCKNO},2,3) <> 'NPN' AND {rankfle.invclass}=" & vINVCLASS & " and isnull({rankfle.subinvclas})= true AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
        Else
            rptRanks.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptRanks.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "SlowMoving.rpt", "MID({rankfle.STOCKNO},2,3) <> 'NPN' AND {rankfle.invclass}=" & vINVCLASS & " and {rankfle.subinvclas}=" & vSUBINVCLAS & " AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
            rptRanks.ReportTitle = Replace(Replace("RANK " & vINVCLASS & vSUBINVCLAS & " SLOW MOVING ACCESSORIES WITH SELLING PRICE", "NULL", ""), "'", "")
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "SlowMoving_A.rpt", "MID({rankfle.STOCKNO},2,3) <> 'NPN' AND {rankfle.invclass}=" & vINVCLASS & " and {rankfle.subinvclas}=" & vSUBINVCLAS & " AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
            rptRanks.ReportTitle = Replace(Replace("RANK " & vINVCLASS & vSUBINVCLAS & "  SLOW MOVING MATERIALS WITH SELLING PRICE", "NULL", ""), "'", "")
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "SlowMoving_M.rpt", "MID({rankfle.STOCKNO},2,3) <> 'NPN' AND {rankfle.invclass}=" & vINVCLASS & " and {rankfle.subinvclas}=" & vSUBINVCLAS & " AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
            rptRanks.ReportTitle = Replace(Replace("RANK " & vINVCLASS & vSUBINVCLAS & "  SLOW MOVING PARTS,ACCESSORIES,MATERIALS WITH SELLING PRICE", "NULL", ""), "'", "")
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "SlowMoving_All.rpt", "MID({rankfle.STOCKNO},2,3) <> 'NPN' AND {rankfle.invclass}=" & vINVCLASS & " and {rankfle.subinvclas}=" & vSUBINVCLAS & " AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
        End If
    Else
        If vSUBINVCLAS = "NULL" Then
            rptRanks.WindowTitle = "Slow Moving Parts for Disposal"
            rptRanks.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptRanks.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "SlowMoving.rpt", "{rankfle.invclass}=" & vINVCLASS & " and isnull({rankfle.subinvclas})= true AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
            rptRanks.ReportTitle = Replace(Replace("RANK " & vINVCLASS & vSUBINVCLAS & " SLOW MOVING ACCESSORIES WITH SELLING PRICE", "NULL", ""), "'", "")
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "SlowMoving_A.rpt", "{rankfle.invclass}=" & vINVCLASS & " and isnull({rankfle.subinvclas})= true AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
            rptRanks.ReportTitle = Replace(Replace("RANK " & vINVCLASS & vSUBINVCLAS & "  SLOW MOVING MATERIALS WITH SELLING PRICE", "NULL", ""), "'", "")
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "SlowMoving_M.rpt", "{rankfle.invclass}=" & vINVCLASS & " and isnull({rankfle.subinvclas})= true AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
            rptRanks.ReportTitle = Replace(Replace("RANK " & vINVCLASS & vSUBINVCLAS & "  SLOW MOVING PARTS,ACCESSORIES,MATERIALS WITH SELLING PRICE", "NULL", ""), "'", "")
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "SlowMoving_All.rpt", "{rankfle.invclass}=" & vINVCLASS & " and isnull({rankfle.subinvclas})= true AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
        Else
            rptRanks.WindowTitle = "Slow Moving Parts for Disposal"
            rptRanks.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptRanks.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "SlowMoving.rpt", "{rankfle.invclass}=" & vINVCLASS & " and {rankfle.subinvclas}=" & vSUBINVCLAS & " AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
            rptRanks.ReportTitle = Replace(Replace("RANK " & vINVCLASS & vSUBINVCLAS & " SLOW MOVING ACCESSORIES WITH SELLING PRICE", "NULL", ""), "'", "")
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "SlowMoving_A.rpt", "{rankfle.invclass}=" & vINVCLASS & " and {rankfle.subinvclas}=" & vSUBINVCLAS & " AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
            rptRanks.ReportTitle = Replace(Replace("RANK " & vINVCLASS & vSUBINVCLAS & "  SLOW MOVING MATERIALS WITH SELLING PRICE", "NULL", ""), "'", "")
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "SlowMoving_M.rpt", "{rankfle.invclass}=" & vINVCLASS & " and {rankfle.subinvclas}=" & vSUBINVCLAS & " AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
            rptRanks.ReportTitle = Replace(Replace("RANK " & vINVCLASS & vSUBINVCLAS & "  SLOW MOVING PARTS,ACCESSORIES,MATERIALS WITH SELLING PRICE", "NULL", ""), "'", "")
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "SlowMoving_All.rpt", "{rankfle.invclass}=" & vINVCLASS & " and {rankfle.subinvclas}=" & vSUBINVCLAS & " AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
        End If
    End If
    LogAudit "V", "INVENTORY RANKING " & vINVCLASS & "-" & vSUBINVCLAS
    Screen.MousePointer = 0
End Sub

Private Sub cmdD_Click()
    PrintInventoryRank "D", ""
End Sub

Private Sub cmdE1_Click()
    PrintInventoryRank "E", "1"
End Sub

Private Sub cmdE2_Click()
    PrintInventoryRank "E", "2"
End Sub

Private Sub cmdE3_Click()
    PrintInventoryRank "E", "3"
End Sub

Private Sub cmdE4_Click()
    PrintInventoryRank "E", "4"
End Sub

Private Sub cmdE5_Click()
    PrintInventoryRank "E", "5"
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Set rsRANKFLE = New ADODB.Recordset
    rsRANKFLE.Open "select date_gen from PMIS_RankFle group by date_gen order by date_gen desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsRANKFLE.EOF And Not rsRANKFLE.BOF Then
        cboDate_Gen.Clear
        Do While Not rsRANKFLE.EOF
            cboDate_Gen.AddItem Null2Date(rsRANKFLE!DATE_GEN)
            rsRANKFLE.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISReports_SlowMoving = Nothing
    UnloadForm Me
End Sub


VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmPMISReports_MoveCat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MOVEMENT CATEGORY"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MoveCat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   4470
   Begin VB.CheckBox chkDontIncludeLocal 
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
      Left            =   210
      TabIndex        =   7
      Top             =   6240
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select date from the list"
      Top             =   90
      Width           =   2205
   End
   Begin Crystal.CrystalReport rptRanks 
      Left            =   5670
      Top             =   6330
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
   Begin wizButton.cmd cmdE1 
      Height          =   405
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "View detailed lift of Rank E1 parts"
      Top             =   3240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   714
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
      MICON           =   "MoveCat.frx":0E42
   End
   Begin wizButton.cmd cmdE2 
      Height          =   405
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "View detailed list of Rank E2 parts"
      Top             =   3720
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   714
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
      MICON           =   "MoveCat.frx":115C
   End
   Begin wizButton.cmd cmdE3 
      Height          =   405
      Left            =   60
      TabIndex        =   3
      ToolTipText     =   "View detailed list of Rank E3 parts"
      Top             =   4200
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   714
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
      MICON           =   "MoveCat.frx":1476
   End
   Begin wizButton.cmd cmdE4 
      Height          =   405
      Left            =   60
      TabIndex        =   4
      ToolTipText     =   "View detailed list of Rank E4 parts"
      Top             =   4680
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   714
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
      MICON           =   "MoveCat.frx":1790
   End
   Begin wizButton.cmd cmdE5 
      Height          =   405
      Left            =   60
      TabIndex        =   5
      ToolTipText     =   "View detailed list of Rank E5 parts"
      Top             =   5160
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   714
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
      MICON           =   "MoveCat.frx":1AAA
   End
   Begin wizButton.cmd cmdA1 
      Height          =   405
      Left            =   60
      TabIndex        =   8
      ToolTipText     =   "View detailed list of Rank A1 parts"
      Top             =   510
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   714
      TX              =   "RANK A1 - SUPER  FAST MOVING PARTS"
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
      MICON           =   "MoveCat.frx":1DC4
   End
   Begin wizButton.cmd cmdA2 
      Height          =   405
      Left            =   60
      TabIndex        =   9
      ToolTipText     =   "View detailed list of Rank A2 parts"
      Top             =   962
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   714
      TX              =   "RANK A2 - SUPER  FAST MOVING PARTS"
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
      MICON           =   "MoveCat.frx":20DE
   End
   Begin wizButton.cmd cmdA3 
      Height          =   405
      Left            =   60
      TabIndex        =   10
      ToolTipText     =   "View detailed list of rank A3 parts"
      Top             =   1414
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   714
      TX              =   "RANK A3 - SUPER  FAST MOVING PARTS"
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
      MICON           =   "MoveCat.frx":23F8
   End
   Begin wizButton.cmd cmdB 
      Height          =   405
      Left            =   60
      TabIndex        =   11
      ToolTipText     =   "View detailed list of Rank C parts"
      Top             =   2318
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   714
      TX              =   "RANK C - MEDIUM MOVING PARTS"
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
      MICON           =   "MoveCat.frx":2712
   End
   Begin wizButton.cmd cmdC 
      Height          =   405
      Left            =   60
      TabIndex        =   12
      ToolTipText     =   "View detailed list of Rank D parts"
      Top             =   2770
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   714
      TX              =   "RANK D  - SLOW MOVING PARTS"
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
      MICON           =   "MoveCat.frx":2A2C
   End
   Begin wizButton.cmd cmdD 
      Height          =   405
      Left            =   6480
      TabIndex        =   13
      ToolTipText     =   "View detailed list of Rank E parts"
      Top             =   1680
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   714
      TX              =   "RANK E  - NON MOVING PARTS"
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
      MICON           =   "MoveCat.frx":2D46
   End
   Begin wizButton.cmd cmdF 
      Height          =   405
      Left            =   60
      TabIndex        =   14
      ToolTipText     =   "View detailed list of new parts"
      Top             =   5640
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   714
      TX              =   "RANK F  - NEW ITEMS"
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
      MICON           =   "MoveCat.frx":3060
   End
   Begin wizButton.cmd cmd_rankA4 
      Height          =   405
      Left            =   60
      TabIndex        =   15
      ToolTipText     =   "View detailed list of rank B parts"
      Top             =   1860
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   714
      TX              =   "RANK B - FAST MOVING PARTS"
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
      MICON           =   "MoveCat.frx":337A
   End
   Begin wizButton.cmd cmd_e2nonmoving 
      Height          =   405
      Left            =   6960
      TabIndex        =   16
      ToolTipText     =   "View detailed list of Rank E2 parts"
      Top             =   2400
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   714
      TX              =   "RANK E2 - NON MOVING PARTS "
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
      MICON           =   "MoveCat.frx":3694
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
      Left            =   690
      TabIndex        =   6
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmPMISReports_MoveCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsRANKFLE                                          As ADODB.Recordset

Sub PrintInventoryRank(vINVCLASS As String, vSUBINVCLAS As String)

    Dim ReportTitle                                    As String
    ReportTitle = "RANK " & vINVCLASS & vSUBINVCLAS & " INVENTORY RANKING REPORT"
    vINVCLASS = "'" & vINVCLASS & "'"
    If vSUBINVCLAS = "" Then vSUBINVCLAS = "NULL" Else vSUBINVCLAS = "'" & vSUBINVCLAS & "'"
    Screen.MousePointer = 11
    rptRanks.ReportTitle = ReportTitle
    If chkDontIncludeLocal.Value = 1 Then
        If vSUBINVCLAS = "NULL" Then

            rptRanks.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptRanks.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "Ranking.rpt", "MID({rankfle.STOCKNO},2,3) <> 'NPN' AND {rankfle.invclass}=" & vINVCLASS & " and isnull({rankfle.subinvclas})= true AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
        Else
            rptRanks.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptRanks.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "Ranking.rpt", "MID({rankfle.STOCKNO},2,3) <> 'NPN' AND {rankfle.invclass}=" & vINVCLASS & " and {rankfle.subinvclas}=" & vSUBINVCLAS & " AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
        End If
    Else
        If vSUBINVCLAS = "NULL" Then
            rptRanks.WindowTitle = "Movement Category Report"
            rptRanks.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptRanks.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "Ranking.rpt", "{rankfle.invclass}=" & vINVCLASS & "  AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
            'PrintSQLReport rptRanks, PMIS_REPORT_PATH & "Ranking.rpt", "{rankfle.invclass}=" & vINVCLASS & " and isnull({rankfle.subinvclas})= true AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
        Else
            rptRanks.WindowTitle = "Movement Category Report"
            rptRanks.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptRanks.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptRanks, PMIS_REPORT_PATH & "Ranking.rpt", "{rankfle.invclass}=" & vINVCLASS & " and {rankfle.subinvclas}=" & vSUBINVCLAS & " AND {rankfle.date_gen} = DateTime(" & Year(cboDate_Gen.Text) & "," & Month(cboDate_Gen.Text) & "," & Day(cboDate_Gen.Text) & ")", DMIS_REPORT_Connection, 1
        End If
    End If
    LogAudit "V", "Movement Category"
    Screen.MousePointer = 0
End Sub

Private Sub cmd_e2nonmoving_Click()
    If IsDate(cboDate_Gen) = False Then
        ShowIsRequiredMsg "Valid Date"
        On Error Resume Next
        cboDate_Gen.SetFocus
        Exit Sub
    End If

    PrintInventoryRank "E", "2"
End Sub

Private Sub cmd_rankA4_Click()
    If IsDate(cboDate_Gen) = False Then
        ShowIsRequiredMsg "Valid Date"
        On Error Resume Next
        cboDate_Gen.SetFocus
        Exit Sub
    End If

    PrintInventoryRank "B", ""
End Sub

Private Sub cmdA1_Click()
    If IsDate(cboDate_Gen) = False Then
        ShowIsRequiredMsg "Valid Date"
        On Error Resume Next
        cboDate_Gen.SetFocus
        Exit Sub
    End If
    PrintInventoryRank "A", "1"
End Sub

Private Sub cmdA2_Click()
    If IsDate(cboDate_Gen) = False Then
        ShowIsRequiredMsg "Valid Date"
        On Error Resume Next
        cboDate_Gen.SetFocus
        Exit Sub
    End If

    PrintInventoryRank "A", "2"
End Sub

Private Sub cmdA3_Click()
    If IsDate(cboDate_Gen) = False Then
        ShowIsRequiredMsg "Valid Date"
        On Error Resume Next
        cboDate_Gen.SetFocus
        Exit Sub
    End If

    PrintInventoryRank "A", "3"
End Sub

Private Sub cmdB_Click()
    If IsDate(cboDate_Gen) = False Then
        ShowIsRequiredMsg "Valid Date"
        On Error Resume Next
        cboDate_Gen.SetFocus
        Exit Sub
    End If

    PrintInventoryRank "C", ""
End Sub

Private Sub cmdC_Click()
    If IsDate(cboDate_Gen) = False Then
        ShowIsRequiredMsg "Valid Date"
        On Error Resume Next
        cboDate_Gen.SetFocus
        Exit Sub
    End If

    PrintInventoryRank "D", ""
End Sub

Private Sub cmdD_Click()
    If IsDate(cboDate_Gen) = False Then
        ShowIsRequiredMsg "Valid Date"
        On Error Resume Next
        cboDate_Gen.SetFocus
        Exit Sub
    End If

    PrintInventoryRank "E", ""
End Sub

Private Sub cmdE1_Click()
    If IsDate(cboDate_Gen) = False Then
        ShowIsRequiredMsg "Valid Date"
        On Error Resume Next
        cboDate_Gen.SetFocus
        Exit Sub
    End If

    PrintInventoryRank "E", "1"
End Sub

Private Sub cmdE2_Click()
    If IsDate(cboDate_Gen) = False Then
        ShowIsRequiredMsg "Valid Date"
        On Error Resume Next
        cboDate_Gen.SetFocus
        Exit Sub
    End If

    PrintInventoryRank "E", "2"
End Sub

Private Sub cmdE3_Click()
    If IsDate(cboDate_Gen) = False Then
        ShowIsRequiredMsg "Valid Date"
        On Error Resume Next
        cboDate_Gen.SetFocus
        Exit Sub
    End If

    PrintInventoryRank "E", "3"
End Sub

Private Sub cmdE4_Click()
    If IsDate(cboDate_Gen) = False Then
        ShowIsRequiredMsg "Valid Date"
        On Error Resume Next
        cboDate_Gen.SetFocus
        Exit Sub
    End If

    PrintInventoryRank "E", "4"
End Sub

Private Sub cmdE5_Click()
    If IsDate(cboDate_Gen) = False Then
        ShowIsRequiredMsg "Valid Date"
        On Error Resume Next
        cboDate_Gen.SetFocus
        Exit Sub
    End If

    PrintInventoryRank "E", "5"
End Sub

'Private Sub cmdE5_Click()
'    If IsDate(cboDate_Gen) = False Then
'        ShowIsRequiredMsg "Valid Date"
'        On Error Resume Next
'        cboDate_Gen.SetFocus
'        Exit Sub
'    End If
'
'    PrintInventoryRank "E", "5"
'End Sub

Private Sub cmdF_Click()
    If IsDate(cboDate_Gen) = False Then
        ShowIsRequiredMsg "Valid Date"
        On Error Resume Next
        cboDate_Gen.SetFocus
        Exit Sub
    End If

    PrintInventoryRank "F", ""
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
    Set frmPMISReports_MoveCat = Nothing
    UnloadForm Me
End Sub


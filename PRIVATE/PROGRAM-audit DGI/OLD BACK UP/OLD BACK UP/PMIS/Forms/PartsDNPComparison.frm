VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmPMISInquiry_PartsDNPComparison 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PARTS DNP COMPARISON"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
   ForeColor       =   &H00DEDFDE&
   Icon            =   "PartsDNPComparison.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   11655
   Begin VB.ComboBox cboRank 
      Height          =   315
      Left            =   10470
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1125
   End
   Begin wizButton.cmd cmdQuery 
      Height          =   525
      Left            =   8880
      TabIndex        =   7
      ToolTipText     =   "Process specified query"
      Top             =   2400
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   926
      TX              =   "Process Query"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "PartsDNPComparison.frx":030A
   End
   Begin Crystal.CrystalReport rptDNPQuery 
      Left            =   11100
      Top             =   4740
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CheckBox chkWStock 
      Caption         =   "with Stock Only"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8880
      TabIndex        =   6
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   8880
      TabIndex        =   15
      ToolTipText     =   "Select one option from list"
      Top             =   420
      Width           =   2685
      Begin VB.OptionButton Opt1 
         Caption         =   "LIST ALL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   30
         TabIndex        =   2
         Top             =   270
         Value           =   -1  'True
         Width           =   2595
      End
      Begin VB.OptionButton Opt4 
         Caption         =   "HARI DNP < DEALER DNP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   30
         TabIndex        =   5
         Top             =   1140
         Width           =   2595
      End
      Begin VB.OptionButton Opt3 
         Caption         =   "HARI DNP = DEALER DNP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   30
         TabIndex        =   4
         Top             =   840
         Width           =   2595
      End
      Begin VB.OptionButton Opt2 
         Caption         =   "HARI DNP > DEALER DNP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   30
         TabIndex        =   3
         Top             =   540
         Width           =   2595
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdQUERY 
      Height          =   5625
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9922
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin wizButton.cmd cmdSearchPartNo 
      Height          =   525
      Left            =   8880
      TabIndex        =   8
      ToolTipText     =   "Search for a part number"
      Top             =   2970
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   926
      TX              =   "Search Part Number"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "PartsDNPComparison.frx":0624
   End
   Begin wizButton.cmd cmdPrintQuery 
      Height          =   525
      Left            =   8880
      TabIndex        =   9
      ToolTipText     =   "Print current query"
      Top             =   3540
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   926
      TX              =   "Print this Query"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "PartsDNPComparison.frx":093E
   End
   Begin wizButton.cmd cmdUpdate 
      Height          =   525
      Left            =   8880
      TabIndex        =   10
      ToolTipText     =   "Update Master File"
      Top             =   4110
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   926
      TX              =   "Update Dealer Master File"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "PartsDNPComparison.frx":0C58
   End
   Begin wizButton.cmd cmdExit 
      Height          =   525
      Left            =   8880
      TabIndex        =   11
      ToolTipText     =   "Exit Window"
      Top             =   4680
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   926
      TX              =   "Exit this Query"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   99
      MICON           =   "PartsDNPComparison.frx":0F72
   End
   Begin wizProgBar.Prg progCPB 
      Height          =   345
      Left            =   9030
      TabIndex        =   17
      ToolTipText     =   "Process progress"
      Top             =   5730
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   609
      Picture         =   "PartsDNPComparison.frx":128C
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "PartsDNPComparison.frx":12A8
      ShowText        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory Rank"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8880
      TabIndex        =   18
      Top             =   150
      Width           =   1845
   End
   Begin VB.Label labCPB 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   8910
      TabIndex        =   16
      Top             =   5250
      Width           =   2625
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "HARI PARTS MASTERFILE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   4260
      TabIndex        =   14
      Top             =   120
      Width           =   4185
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "DEALER PARTS MASTERFILE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   90
      TabIndex        =   13
      Top             =   120
      Width           =   4125
   End
   Begin VB.Label labAydi 
      Caption         =   "Label1"
      Height          =   195
      Left            =   7440
      TabIndex        =   12
      Top             =   4890
      Visible         =   0   'False
      Width           =   30
   End
End
Attribute VB_Name = "frmPMISInquiry_PartsDNPComparison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPartmasDNPP                                                     As ADODB.Recordset
Dim KCNT                                                              As Integer

Function PARTSINQUIRYBFound(ByVal str2find) As Boolean
    On Error GoTo BFoundErr
    Dim result                                                        As Boolean
    Dim rsBClone                                                      As ADODB.Recordset
    result = False
    If Not IsNull(str2find) Then
        Set rsBClone = New ADODB.Recordset
        Set rsBClone = rsPartmasDNPP.Clone

        rsBClone.Find "PARTNO = '" & str2find & "'"
        result = Not rsBClone.EOF
        If result Then
            rsPartmasDNPP.Bookmark = rsBClone.Bookmark
        End If
        Set rsBClone = Nothing
    End If
    PARTSINQUIRYBFound = result
    Exit Function
BFoundErr:
    ShowVBError
End Function

Sub initPARTSINQUIRYGrid()
    With grdQUERY
        .Row = 0
        .FormatString = "Part Number  | Part Desc               |    DNP     | " & _
                        "Part Number  | Part Desc               |    DNP     "
    End With
End Sub

Sub FillPARTSINQUIRYGrid()

End Sub
'    kcnt = 0
'        Screen.MousePointer = 11
'        rsPartmasDNPP.MoveFirst
'        cmdSearchPartNo.Enabled = False
'        cmdExit.Enabled = False
'        Do While Not rsPartmasDNPP.EOF
'            kcnt = kcnt + 1
'            grdQUERY.AddItem Null2String(rsPartmasDNPP!partno) & Chr(9) & _
             '                             Null2String(rsPartmasDNPP!PartDesc) & Chr(9) & _
             '                             ToDoubleNumber(Null2String(rsPartmasDNPP!DEALERDNP)) & Chr(9) & _
             '                             Null2String(rsPartmasDNPP!PARTNUMBER) & Chr(9) & _
             '                             Null2String(rsPartmasDNPP!DESCRIPTIO) & Chr(9) & _
             '                             ToDoubleNumber(N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP))
'            rsPartmasDNPP.MoveNext
'            If kcnt = 1 Then grdQUERY.RemoveItem 1
'            progCPB.Value = (kcnt / rsPartmasDNPP.RecordCount) * 100
'            labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & kcnt
'            DoEvents
'            DoEvents
'        Loop
'        LogAudit "V", "Parts DNP Comparision"
'        cmdSearchPartNo.Enabled = True
'        cmdExit.Enabled = True
'        Screen.MousePointer = 0
'
'    Exit Sub
'
'ErrorCode:
'    ShowVBError
'    Exit Sub
'End Sub

Private Sub cmdPrintQuery_Click()
    If cboRank.Text = "ALL" Then
        If opt1.Value = True Then
            rptDNPQuery.ReportTitle = "LIST ALL"
            If chkWStock.Value = 1 Then
                rptDNPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptDNPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptDNPQuery, PMIS_REPORT_PATH & "DNPcomparison.rpt", "{PMIS_PARTMAS.ONHAND} > 0", DMIS_REPORT_Connection, 1
            Else
                rptDNPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptDNPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptDNPQuery, PMIS_REPORT_PATH & "DNPcomparison.rpt", "", DMIS_REPORT_Connection, 1
            End If
        End If
        If opt2.Value = True Then
            rptDNPQuery.ReportTitle = "HARI DNP > DEALER DNP"
            If chkWStock.Value = 1 Then
                rptDNPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptDNPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptDNPQuery, PMIS_REPORT_PATH & "DNPcomparison.rpt", "ROUND({PMIS_DNPP.DNPP},2) > ROUND(({PMIS_PARTMAS.MAC} * " & ConvertToBIRDecimalFormat(VAT_RATE) & "),2) AND {PMIS_PARTMAS.ONHAND} > 0", DMIS_REPORT_Connection, 1
            Else
                rptDNPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptDNPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptDNPQuery, PMIS_REPORT_PATH & "DNPcomparison.rpt", "ROUND({PMIS_DNPP.DNPP},2) > ROUND(({PMIS_PARTMAS.MAC} * " & ConvertToBIRDecimalFormat(VAT_RATE) & "),2)", DMIS_REPORT_Connection, 1
            End If
        End If
        If Opt3.Value = True Then
            rptDNPQuery.ReportTitle = "HARI DNP = DEALER DNP"
            If chkWStock.Value = 1 Then
                rptDNPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptDNPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptDNPQuery, PMIS_REPORT_PATH & "DNPcomparison.rpt", "ROUND({PMIS_DNPP.DNPP},2) = ROUND(({PMIS_PARTMAS.MAC} * " & ConvertToBIRDecimalFormat(VAT_RATE) & " ),2) AND {PMIS_PARTMAS.ONHAND} > 0", DMIS_REPORT_Connection, 1
            Else
                rptDNPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptDNPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptDNPQuery, PMIS_REPORT_PATH & "DNPcomparison.rpt", "ROUND({PMIS_DNPP.DNPP},2) = ROUND(({PMIS_PARTMAS.MAC} * " & ConvertToBIRDecimalFormat(VAT_RATE) & " ),2)", DMIS_REPORT_Connection, 1
            End If
        End If
        If Opt4.Value = True Then
            rptDNPQuery.ReportTitle = "HARI DNP < DEALER DNP"
            If chkWStock.Value = 1 Then
                rptDNPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptDNPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptDNPQuery, PMIS_REPORT_PATH & "DNPcomparison.rpt", "ROUND({PMIS_DNPP.DNPP},2) < ROUND(({PMIS_PARTMAS.MAC} * " & ConvertToBIRDecimalFormat(VAT_RATE) & " ),2) AND {PMIS_PARTMAS.ONHAND} > 0", DMIS_REPORT_Connection, 1
            Else
                rptDNPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptDNPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptDNPQuery, PMIS_REPORT_PATH & "DNPcomparison.rpt", "ROUND({PMIS_DNPP.DNPP},2) < ROUND(({PMIS_PARTMAS.MAC} * " & ConvertToBIRDecimalFormat(VAT_RATE) & " ),2)", DMIS_REPORT_Connection, 1
            End If
        End If
    Else
        Dim InvClass, SubInvClas                                      As String
        If cboRank.Text = "A1" Then
            InvClass = "'A'": SubInvClas = "{PMIS_PARTMAS.SUBINVCLAS} = '1'"
        End If
        If cboRank.Text = "A2" Then
            InvClass = "'A'": SubInvClas = "{PMIS_PARTMAS.SUBINVCLAS} = '2'"
        End If
        If cboRank.Text = "A3" Then
            InvClass = "'A'": SubInvClas = "{PMIS_PARTMAS.SUBINVCLAS} = '3'"
        End If
        If cboRank.Text = "B" Then
            InvClass = "'B'": SubInvClas = "ISNULL({PMIS_PARTMAS.SUBINVCLAS}) = true"
        End If
        If cboRank.Text = "C" Then
            InvClass = "'C'": SubInvClas = "ISNULL({PMIS_PARTMAS.SUBINVCLAS}) = true"
        End If
        If cboRank.Text = "D" Then
            InvClass = "'D'": SubInvClas = "ISNULL({PMIS_PARTMAS.SUBINVCLAS}) = true"
        End If
        If cboRank.Text = "E1" Then
            InvClass = "'E'": SubInvClas = "{PMIS_PARTMAS.SUBINVCLAS} = '1'"
        End If
        If cboRank.Text = "E2" Then
            InvClass = "'E'": SubInvClas = "{PMIS_PARTMAS.SUBINVCLAS} = '2'"
        End If
        If cboRank.Text = "E3" Then
            InvClass = "'E'": SubInvClas = "{PMIS_PARTMAS.SUBINVCLAS} = '3'"
        End If
        If cboRank.Text = "E4" Then
            InvClass = "'E'": SubInvClas = "{PMIS_PARTMAS.SUBINVCLAS} = '4'"
        End If
        If cboRank.Text = "E5" Then
            InvClass = "'E'": SubInvClas = "{PMIS_PARTMAS.SUBINVCLAS} = '5'"
        End If
        If cboRank.Text = "F" Then
            InvClass = "'F'": SubInvClas = "ISNULL({PMIS_PARTMAS.SUBINVCLAS}) = true"
        End If
        If opt1.Value = True Then
            rptDNPQuery.ReportTitle = "LIST ALL"
            If chkWStock.Value = 1 Then
                rptDNPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptDNPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptDNPQuery, PMIS_REPORT_PATH & "DNPcomparison.rpt", "{PMIS_PARTMAS.ONHAND} > 0 AND ({PMIS_PARTMAS.INVCLASS} = " & InvClass & " AND " & SubInvClas & ")", DMIS_REPORT_Connection, 1
            Else
                rptDNPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptDNPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptDNPQuery, PMIS_REPORT_PATH & "DNPcomparison.rpt", "({PMIS_PARTMAS.INVCLASS} = " & InvClass & " AND " & SubInvClas & ")", DMIS_REPORT_Connection, 1
            End If
        End If
        If opt2.Value = True Then
            rptDNPQuery.ReportTitle = "HARI DNP > DEALER DNP"
            If chkWStock.Value = 1 Then
                rptDNPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptDNPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptDNPQuery, PMIS_REPORT_PATH & "DNPcomparison.rpt", "ROUND({PMIS_DNPP.DNPP},2) > ROUND(({PMIS_PARTMAS.MAC} * ConvertToBIRDecimalFormat(VAT_RATE)),2) AND {PMIS_PARTMAS.ONHAND} > 0  AND ({PMIS_PARTMAS.INVCLASS} = " & InvClass & " AND " & SubInvClas & ")", DMIS_REPORT_Connection, 1
            Else
                rptDNPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptDNPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptDNPQuery, PMIS_REPORT_PATH & "DNPcomparison.rpt", "ROUND({PMIS_DNPP.DNPP},2) > ROUND(({PMIS_PARTMAS.MAC} * ConvertToBIRDecimalFormat(VAT_RATE)),2) AND ({PMIS_PARTMAS.INVCLASS} = " & InvClass & " AND " & SubInvClas & ")", DMIS_REPORT_Connection, 1
            End If
        End If
        If Opt3.Value = True Then
            rptDNPQuery.ReportTitle = "HARI DNP = DEALER DNP"
            If chkWStock.Value = 1 Then
                rptDNPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptDNPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptDNPQuery, PMIS_REPORT_PATH & "DNPcomparison.rpt", "ROUND({PMIS_DNPP.DNPP},2) = ROUND(({PMIS_PARTMAS.MAC} * ConvertToBIRDecimalFormat(VAT_RATE)),2) AND {PMIS_PARTMAS.ONHAND} > 0 AND ({PMIS_PARTMAS.INVCLASS} = " & InvClass & " AND " & SubInvClas & ")", DMIS_REPORT_Connection, 1
            Else
                rptDNPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptDNPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptDNPQuery, PMIS_REPORT_PATH & "DNPcomparison.rpt", "ROUND({PMIS_DNPP.DNPP},2) = ROUND(({PMIS_PARTMAS.MAC} * ConvertToBIRDecimalFormat(VAT_RATE)),2) AND ({PMIS_PARTMAS.INVCLASS} = " & InvClass & " AND " & SubInvClas & ")", DMIS_REPORT_Connection, 1
            End If
        End If
        If Opt4.Value = True Then
            rptDNPQuery.ReportTitle = "HARI DNP < DEALER DNP"
            If chkWStock.Value = 1 Then
                rptDNPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptDNPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptDNPQuery, PMIS_REPORT_PATH & "DNPcomparison.rpt", "ROUND({PMIS_DNPP.DNPP},2) < ROUND(({PMIS_PARTMAS.MAC} * ConvertToBIRDecimalFormat(VAT_RATE)),2) AND {PMIS_PARTMAS.ONHAND} > 0 AND ({PMIS_PARTMAS.INVCLASS} = " & InvClass & " AND " & SubInvClas & ")", DMIS_REPORT_Connection, 1
            Else
                rptDNPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptDNPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptDNPQuery, PMIS_REPORT_PATH & "DNPcomparison.rpt", "ROUND({PMIS_DNPP.DNPP},2) < ROUND(({PMIS_PARTMAS.MAC} * ConvertToBIRDecimalFormat(VAT_RATE)),2) AND ({PMIS_PARTMAS.INVCLASS} = " & InvClass & " AND " & SubInvClas & ")", DMIS_REPORT_Connection, 1
            End If
        End If
    End If

    LogAudit "V", "Parts DNP Comparision"
End Sub

Private Sub cmdQuery_Click()
    Dim rsPartmasDNPP                                                 As ADODB.Recordset
    If cboRank.Text = "ALL" Then
        Screen.MousePointer = 11
        cleargrid grdQUERY
        If opt1.Value = True Then
            If chkWStock.Value = 1 Then
                Set rsPartmasDNPP = New ADODB.Recordset
                'rsPartmasDNPP.Open "Select PMIS_PARTMAS.ONHAND,PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.DNPP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where PMIS_PARTMAS.ONHAND > 0 order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                rsPartmasDNPP.Open "Select PMIS_PARTMAS.ONHAND,PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC, ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP , PMIS_DNPP.PARTNUMBER , PMIS_DNPP.DESCRIPTIO, ROUND(PMIS_DNPP.DNPP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.DNPP,2) > ROUND(PMIS_PARTMAS.DNP,2) AND PMIS_PARTMAS.ONHAND > 0 order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                initPARTSINQUIRYGrid
                If Not rsPartmasDNPP.EOF And Not rsPartmasDNPP.BOF Then
                    KCNT = 0
                    Screen.MousePointer = 11
                    rsPartmasDNPP.MoveFirst
                    cmdSearchPartNo.Enabled = False
                    cmdExit.Enabled = False
                    Do While Not rsPartmasDNPP.EOF
                        KCNT = KCNT + 1
                        grdQUERY.AddItem Null2String(rsPartmasDNPP!PARTNO) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTDESC) & Chr(9) & _
                                         ToDoubleNumber(Null2String(rsPartmasDNPP!DEALERDNP)) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTNUMBER) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!DESCRIPTIO) & Chr(9) & _
                                         ToDoubleNumber(N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP))
                        rsPartmasDNPP.MoveNext
                        If KCNT = 1 Then grdQUERY.RemoveItem 1
                        progCPB.Value = (KCNT / rsPartmasDNPP.RecordCount) * 100
                        labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
                        DoEvents
                        DoEvents
                    Loop
                    LogAudit "V", "Parts DNP Comparision"
                    cmdSearchPartNo.Enabled = True
                    cmdExit.Enabled = True
                    Screen.MousePointer = 0
                Else
                    cleargrid grdQUERY
                End If
            Else
                Set rsPartmasDNPP = New ADODB.Recordset
                rsPartmasDNPP.Open "Select PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.DNPP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                initPARTSINQUIRYGrid
                If Not rsPartmasDNPP.EOF And Not rsPartmasDNPP.BOF Then
                    KCNT = 0
                    Screen.MousePointer = 11
                    rsPartmasDNPP.MoveFirst
                    cmdSearchPartNo.Enabled = False
                    cmdExit.Enabled = False
                    Do While Not rsPartmasDNPP.EOF
                        KCNT = KCNT + 1
                        grdQUERY.AddItem Null2String(rsPartmasDNPP!PARTNO) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTDESC) & Chr(9) & _
                                         ToDoubleNumber(Null2String(rsPartmasDNPP!DEALERDNP)) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTNUMBER) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!DESCRIPTIO) & Chr(9) & _
                                         ToDoubleNumber(N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP))
                        rsPartmasDNPP.MoveNext
                        If KCNT = 1 Then grdQUERY.RemoveItem 1
                        progCPB.Value = (KCNT / rsPartmasDNPP.RecordCount) * 100
                        labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
                        DoEvents
                        DoEvents
                    Loop
                    LogAudit "V", "Parts DNP Comparision"
                    cmdSearchPartNo.Enabled = True
                    cmdExit.Enabled = True
                    Screen.MousePointer = 0
                Else
                    cleargrid grdQUERY
                End If
            End If
        End If
        If opt2.Value = True Then
            If chkWStock.Value = 1 Then
                Set rsPartmasDNPP = New ADODB.Recordset
                rsPartmasDNPP.Open "Select PMIS_PARTMAS.ONHAND,PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.DNPP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.DNPP,2) > ROUND(PMIS_PARTMAS.DNP,2) AND PMIS_PARTMAS.ONHAND > 0 order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                initPARTSINQUIRYGrid
                If Not rsPartmasDNPP.EOF And Not rsPartmasDNPP.BOF Then
                    KCNT = 0
                    Screen.MousePointer = 11
                    rsPartmasDNPP.MoveFirst
                    cmdSearchPartNo.Enabled = False
                    cmdExit.Enabled = False
                    Do While Not rsPartmasDNPP.EOF
                        KCNT = KCNT + 1
                        grdQUERY.AddItem Null2String(rsPartmasDNPP!PARTNO) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTDESC) & Chr(9) & _
                                         ToDoubleNumber(Null2String(rsPartmasDNPP!DEALERDNP)) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTNUMBER) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!DESCRIPTIO) & Chr(9) & _
                                         ToDoubleNumber(N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP))
                        rsPartmasDNPP.MoveNext
                        If KCNT = 1 Then grdQUERY.RemoveItem 1
                        progCPB.Value = (KCNT / rsPartmasDNPP.RecordCount) * 100
                        labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
                        DoEvents
                        DoEvents
                    Loop
                    LogAudit "V", "Parts DNP Comparision"
                    cmdSearchPartNo.Enabled = True
                    cmdExit.Enabled = True
                    Screen.MousePointer = 0
                Else
                    cleargrid grdQUERY
                End If
            Else
                Set rsPartmasDNPP = New ADODB.Recordset
                rsPartmasDNPP.Open "Select PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.SRP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.SRP,2) > ROUND(PMIS_PARTMAS.DNP,2) order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                initPARTSINQUIRYGrid
                If Not rsPartmasDNPP.EOF And Not rsPartmasDNPP.BOF Then
                    KCNT = 0
                    Screen.MousePointer = 11
                    rsPartmasDNPP.MoveFirst
                    cmdSearchPartNo.Enabled = False
                    cmdExit.Enabled = False
                    Do While Not rsPartmasDNPP.EOF
                        KCNT = KCNT + 1
                        grdQUERY.AddItem Null2String(rsPartmasDNPP!PARTNO) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTDESC) & Chr(9) & _
                                         ToDoubleNumber(Null2String(rsPartmasDNPP!DEALERDNP)) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTNUMBER) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!DESCRIPTIO) & Chr(9) & _
                                         ToDoubleNumber(N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP))
                        rsPartmasDNPP.MoveNext
                        If KCNT = 1 Then grdQUERY.RemoveItem 1
                        progCPB.Value = (KCNT / rsPartmasDNPP.RecordCount) * 100
                        labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
                        DoEvents
                        DoEvents
                    Loop
                    LogAudit "V", "Parts DNP Comparision"
                    cmdSearchPartNo.Enabled = True
                    cmdExit.Enabled = True
                    Screen.MousePointer = 0
                Else
                    cleargrid grdQUERY
                End If

            End If
        End If
        If Opt3.Value = True Then
            If chkWStock.Value = 1 Then
                Set rsPartmasDNPP = New ADODB.Recordset
                rsPartmasDNPP.Open "Select PMIS_PARTMAS.ONHAND,PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.DNPP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.DNPP,2) = ROUND(PMIS_PARTMAS.DNP,2) AND PMIS_PARTMAS.ONHAND > 0 order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                initPARTSINQUIRYGrid
                If Not rsPartmasDNPP.EOF And Not rsPartmasDNPP.BOF Then
                    KCNT = 0
                    Screen.MousePointer = 11
                    rsPartmasDNPP.MoveFirst
                    cmdSearchPartNo.Enabled = False
                    cmdExit.Enabled = False
                    Do While Not rsPartmasDNPP.EOF
                        KCNT = KCNT + 1
                        grdQUERY.AddItem Null2String(rsPartmasDNPP!PARTNO) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTDESC) & Chr(9) & _
                                         ToDoubleNumber(Null2String(rsPartmasDNPP!DEALERDNP)) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTNUMBER) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!DESCRIPTIO) & Chr(9) & _
                                         ToDoubleNumber(N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP))
                        rsPartmasDNPP.MoveNext
                        If KCNT = 1 Then grdQUERY.RemoveItem 1
                        progCPB.Value = (KCNT / rsPartmasDNPP.RecordCount) * 100
                        labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
                        DoEvents
                        DoEvents
                    Loop
                    LogAudit "V", "Parts DNP Comparision"
                    cmdSearchPartNo.Enabled = True
                    cmdExit.Enabled = True
                    Screen.MousePointer = 0
                Else
                    cleargrid grdQUERY
                End If

            Else
                Set rsPartmasDNPP = New ADODB.Recordset
                rsPartmasDNPP.Open "Select PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.DNPP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER  where ROUND(PMIS_DNPP.DNPP,2) = ROUND(PMIS_PARTMAS.DNP,2) order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                initPARTSINQUIRYGrid
                If Not rsPartmasDNPP.EOF And Not rsPartmasDNPP.BOF Then
                    KCNT = 0
                    Screen.MousePointer = 11
                    rsPartmasDNPP.MoveFirst
                    cmdSearchPartNo.Enabled = False
                    cmdExit.Enabled = False
                    Do While Not rsPartmasDNPP.EOF
                        KCNT = KCNT + 1
                        grdQUERY.AddItem Null2String(rsPartmasDNPP!PARTNO) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTDESC) & Chr(9) & _
                                         ToDoubleNumber(Null2String(rsPartmasDNPP!DEALERDNP)) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTNUMBER) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!DESCRIPTIO) & Chr(9) & _
                                         ToDoubleNumber(N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP))
                        rsPartmasDNPP.MoveNext
                        If KCNT = 1 Then grdQUERY.RemoveItem 1
                        progCPB.Value = (KCNT / rsPartmasDNPP.RecordCount) * 100
                        labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
                        DoEvents
                        DoEvents
                    Loop
                    LogAudit "V", "Parts DNP Comparision"
                    cmdSearchPartNo.Enabled = True
                    cmdExit.Enabled = True
                    Screen.MousePointer = 0
                Else
                    cleargrid grdQUERY
                End If
            End If
        End If
        If Opt4.Value = True Then
            If chkWStock.Value = 1 Then
                Set rsPartmasDNPP = New ADODB.Recordset
                rsPartmasDNPP.Open "Select PMIS_PARTMAS.ONHAND,PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.DNPP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.DNPP,2) < ROUND(PMIS_PARTMAS.DNP,2) AND PMIS_PARTMAS.ONHAND > 0 order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                initPARTSINQUIRYGrid
                If Not rsPartmasDNPP.EOF And Not rsPartmasDNPP.BOF Then
                    KCNT = 0
                    Screen.MousePointer = 11
                    rsPartmasDNPP.MoveFirst
                    cmdSearchPartNo.Enabled = False
                    cmdExit.Enabled = False
                    Do While Not rsPartmasDNPP.EOF
                        KCNT = KCNT + 1
                        grdQUERY.AddItem Null2String(rsPartmasDNPP!PARTNO) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTDESC) & Chr(9) & _
                                         ToDoubleNumber(Null2String(rsPartmasDNPP!DEALERDNP)) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTNUMBER) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!DESCRIPTIO) & Chr(9) & _
                                         ToDoubleNumber(N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP))
                        rsPartmasDNPP.MoveNext
                        If KCNT = 1 Then grdQUERY.RemoveItem 1
                        progCPB.Value = (KCNT / rsPartmasDNPP.RecordCount) * 100
                        labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
                        DoEvents
                        DoEvents
                    Loop
                    LogAudit "V", "Parts DNP Comparision"
                    cmdSearchPartNo.Enabled = True
                    cmdExit.Enabled = True
                    Screen.MousePointer = 0
                Else
                    cleargrid grdQUERY
                End If
            Else
                Set rsPartmasDNPP = New ADODB.Recordset
                rsPartmasDNPP.Open "Select PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.DNPP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER  where ROUND(PMIS_DNPP.DNPP,2) < ROUND(PMIS_PARTMAS.DNP,2) order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                initPARTSINQUIRYGrid
                If Not rsPartmasDNPP.EOF And Not rsPartmasDNPP.BOF Then
                    KCNT = 0
                    Screen.MousePointer = 11
                    rsPartmasDNPP.MoveFirst
                    cmdSearchPartNo.Enabled = False
                    cmdExit.Enabled = False
                    Do While Not rsPartmasDNPP.EOF
                        KCNT = KCNT + 1
                        grdQUERY.AddItem Null2String(rsPartmasDNPP!PARTNO) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTDESC) & Chr(9) & _
                                         ToDoubleNumber(Null2String(rsPartmasDNPP!DEALERDNP)) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTNUMBER) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!DESCRIPTIO) & Chr(9) & _
                                         ToDoubleNumber(N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP))
                        rsPartmasDNPP.MoveNext
                        If KCNT = 1 Then grdQUERY.RemoveItem 1
                        progCPB.Value = (KCNT / rsPartmasDNPP.RecordCount) * 100
                        labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
                        DoEvents
                        DoEvents
                    Loop
                    LogAudit "V", "Parts DNP Comparision"
                    cmdSearchPartNo.Enabled = True
                    cmdExit.Enabled = True
                    Screen.MousePointer = 0
                Else
                    cleargrid grdQUERY
                End If
            End If
        End If
        Screen.MousePointer = 0

    Else

        Dim InvClass, SubInvClas                                      As String
        If cboRank.Text = "A1" Then
            InvClass = "'A'": SubInvClas = "='1'"
        End If
        If cboRank.Text = "A2" Then
            InvClass = "'A'": SubInvClas = "='2'"
        End If
        If cboRank.Text = "A3" Then
            InvClass = "'A'": SubInvClas = "='3'"
        End If
        If cboRank.Text = "B" Then
            InvClass = "'B'": SubInvClas = "IS NULL"
        End If
        If cboRank.Text = "C" Then
            InvClass = "'C'": SubInvClas = "IS NULL"
        End If
        If cboRank.Text = "D" Then
            InvClass = "'D'": SubInvClas = "IS NULL"
        End If
        If cboRank.Text = "E1" Then
            InvClass = "'E'": SubInvClas = "='1'"
        End If
        If cboRank.Text = "E2" Then
            InvClass = "'E'": SubInvClas = "='2'"
        End If
        If cboRank.Text = "E3" Then
            InvClass = "'E'": SubInvClas = "='3'"
        End If
        If cboRank.Text = "E4" Then
            InvClass = "'E'": SubInvClas = "='4'"
        End If
        If cboRank.Text = "E5" Then
            InvClass = "'E'": SubInvClas = "='5'"
        End If
        If cboRank.Text = "F" Then
            InvClass = "'F'": SubInvClas = "IS NULL"
        End If
        Screen.MousePointer = 11
        cleargrid grdQUERY
        If opt1.Value = True Then
            If chkWStock.Value = 1 Then
                Set rsPartmasDNPP = New ADODB.Recordset
                rsPartmasDNPP.Open "Select PMIS_PARTMAS.ONHAND,PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.DNPP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where PMIS_PARTMAS.ONHAND > 0 order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                initPARTSINQUIRYGrid
                If Not rsPartmasDNPP.EOF And Not rsPartmasDNPP.BOF Then
                    KCNT = 0
                    Screen.MousePointer = 11
                    rsPartmasDNPP.MoveFirst
                    cmdSearchPartNo.Enabled = False
                    cmdExit.Enabled = False
                    Do While Not rsPartmasDNPP.EOF
                        KCNT = KCNT + 1
                        grdQUERY.AddItem Null2String(rsPartmasDNPP!PARTNO) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTDESC) & Chr(9) & _
                                         ToDoubleNumber(Null2String(rsPartmasDNPP!DEALERDNP)) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTNUMBER) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!DESCRIPTIO) & Chr(9) & _
                                         ToDoubleNumber(N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP))
                        rsPartmasDNPP.MoveNext
                        If KCNT = 1 Then grdQUERY.RemoveItem 1
                        progCPB.Value = (KCNT / rsPartmasDNPP.RecordCount) * 100
                        labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
                        DoEvents
                        DoEvents
                    Loop
                    LogAudit "V", "Parts DNP Comparision"
                    cmdSearchPartNo.Enabled = True
                    cmdExit.Enabled = True
                    Screen.MousePointer = 0
                Else
                    cleargrid grdQUERY
                End If
            Else
                Set rsPartmasDNPP = New ADODB.Recordset
                rsPartmasDNPP.Open "Select PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.DNPP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                initPARTSINQUIRYGrid
                If Not rsPartmasDNPP.EOF And Not rsPartmasDNPP.BOF Then
                    KCNT = 0
                    Screen.MousePointer = 11
                    rsPartmasDNPP.MoveFirst
                    cmdSearchPartNo.Enabled = False
                    cmdExit.Enabled = False
                    Do While Not rsPartmasDNPP.EOF
                        KCNT = KCNT + 1
                        grdQUERY.AddItem Null2String(rsPartmasDNPP!PARTNO) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTDESC) & Chr(9) & _
                                         ToDoubleNumber(Null2String(rsPartmasDNPP!DEALERDNP)) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTNUMBER) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!DESCRIPTIO) & Chr(9) & _
                                         ToDoubleNumber(N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP))
                        rsPartmasDNPP.MoveNext
                        If KCNT = 1 Then grdQUERY.RemoveItem 1
                        progCPB.Value = (KCNT / rsPartmasDNPP.RecordCount) * 100
                        labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
                        DoEvents
                        DoEvents
                    Loop
                    LogAudit "V", "Parts DNP Comparision"
                    cmdSearchPartNo.Enabled = True
                    cmdExit.Enabled = True
                    Screen.MousePointer = 0
                Else
                    cleargrid grdQUERY
                End If
            End If
        End If
        If opt2.Value = True Then
            If chkWStock.Value = 1 Then
                Set rsPartmasDNPP = New ADODB.Recordset
                rsPartmasDNPP.Open "Select PMIS_PARTMAS.ONHAND,PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.DNPP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.DNPP,2) > ROUND(PMIS_PARTMAS.DNP,2) AND PMIS_PARTMAS.ONHAND > 0 order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                initPARTSINQUIRYGrid
                If Not rsPartmasDNPP.EOF And Not rsPartmasDNPP.BOF Then
                    KCNT = 0
                    Screen.MousePointer = 11
                    rsPartmasDNPP.MoveFirst
                    cmdSearchPartNo.Enabled = False
                    cmdExit.Enabled = False
                    Do While Not rsPartmasDNPP.EOF
                        KCNT = KCNT + 1
                        grdQUERY.AddItem Null2String(rsPartmasDNPP!PARTNO) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTDESC) & Chr(9) & _
                                         ToDoubleNumber(Null2String(rsPartmasDNPP!DEALERDNP)) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTNUMBER) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!DESCRIPTIO) & Chr(9) & _
                                         ToDoubleNumber(N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP))
                        rsPartmasDNPP.MoveNext
                        If KCNT = 1 Then grdQUERY.RemoveItem 1
                        progCPB.Value = (KCNT / rsPartmasDNPP.RecordCount) * 100
                        labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
                        DoEvents
                        DoEvents
                    Loop
                    LogAudit "V", "Parts DNP Comparision"
                    cmdSearchPartNo.Enabled = True
                    cmdExit.Enabled = True
                    Screen.MousePointer = 0
                Else
                    cleargrid grdQUERY
                End If
            Else
                Set rsPartmasDNPP = New ADODB.Recordset
                rsPartmasDNPP.Open "Select PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.SRP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.SRP,2) > ROUND(PMIS_PARTMAS.DNP,2) order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                initPARTSINQUIRYGrid
                If Not rsPartmasDNPP.EOF And Not rsPartmasDNPP.BOF Then
                    KCNT = 0
                    Screen.MousePointer = 11
                    rsPartmasDNPP.MoveFirst
                    cmdSearchPartNo.Enabled = False
                    cmdExit.Enabled = False
                    Do While Not rsPartmasDNPP.EOF
                        KCNT = KCNT + 1
                        grdQUERY.AddItem Null2String(rsPartmasDNPP!PARTNO) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTDESC) & Chr(9) & _
                                         ToDoubleNumber(Null2String(rsPartmasDNPP!DEALERDNP)) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTNUMBER) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!DESCRIPTIO) & Chr(9) & _
                                         ToDoubleNumber(N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP))
                        rsPartmasDNPP.MoveNext
                        If KCNT = 1 Then grdQUERY.RemoveItem 1
                        progCPB.Value = (KCNT / rsPartmasDNPP.RecordCount) * 100
                        labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
                        DoEvents
                        DoEvents
                    Loop
                    LogAudit "V", "Parts DNP Comparision"
                    cmdSearchPartNo.Enabled = True
                    cmdExit.Enabled = True
                    Screen.MousePointer = 0
                Else
                    cleargrid grdQUERY
                End If
            End If
        End If
        If Opt3.Value = True Then
            If chkWStock.Value = 1 Then
                Set rsPartmasDNPP = New ADODB.Recordset
                rsPartmasDNPP.Open "Select PMIS_PARTMAS.ONHAND,PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.DNPP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.DNPP,2) = ROUND(PMIS_PARTMAS.DNP,2) AND PMIS_PARTMAS.ONHAND > 0 order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                initPARTSINQUIRYGrid
                If Not rsPartmasDNPP.EOF And Not rsPartmasDNPP.BOF Then
                    KCNT = 0
                    Screen.MousePointer = 11
                    rsPartmasDNPP.MoveFirst
                    cmdSearchPartNo.Enabled = False
                    cmdExit.Enabled = False
                    Do While Not rsPartmasDNPP.EOF
                        KCNT = KCNT + 1
                        grdQUERY.AddItem Null2String(rsPartmasDNPP!PARTNO) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTDESC) & Chr(9) & _
                                         ToDoubleNumber(Null2String(rsPartmasDNPP!DEALERDNP)) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTNUMBER) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!DESCRIPTIO) & Chr(9) & _
                                         ToDoubleNumber(N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP))
                        rsPartmasDNPP.MoveNext
                        If KCNT = 1 Then grdQUERY.RemoveItem 1
                        progCPB.Value = (KCNT / rsPartmasDNPP.RecordCount) * 100
                        labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
                        DoEvents
                        DoEvents
                    Loop
                    LogAudit "V", "Parts DNP Comparision"
                    cmdSearchPartNo.Enabled = True
                    cmdExit.Enabled = True
                    Screen.MousePointer = 0
                Else
                    cleargrid grdQUERY
                End If
            Else
                Set rsPartmasDNPP = New ADODB.Recordset
                rsPartmasDNPP.Open "Select PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.DNPP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.DNPP,2) = ROUND(PMIS_PARTMAS.DNP,2) order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                initPARTSINQUIRYGrid
                If Not rsPartmasDNPP.EOF And Not rsPartmasDNPP.BOF Then
                    KCNT = 0
                    Screen.MousePointer = 11
                    rsPartmasDNPP.MoveFirst
                    cmdSearchPartNo.Enabled = False
                    cmdExit.Enabled = False
                    Do While Not rsPartmasDNPP.EOF
                        KCNT = KCNT + 1
                        grdQUERY.AddItem Null2String(rsPartmasDNPP!PARTNO) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTDESC) & Chr(9) & _
                                         ToDoubleNumber(Null2String(rsPartmasDNPP!DEALERDNP)) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTNUMBER) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!DESCRIPTIO) & Chr(9) & _
                                         ToDoubleNumber(N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP))
                        rsPartmasDNPP.MoveNext
                        If KCNT = 1 Then grdQUERY.RemoveItem 1
                        progCPB.Value = (KCNT / rsPartmasDNPP.RecordCount) * 100
                        labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
                        DoEvents
                        DoEvents
                    Loop
                    LogAudit "V", "Parts DNP Comparision"
                    cmdSearchPartNo.Enabled = True
                    cmdExit.Enabled = True
                    Screen.MousePointer = 0
                Else
                    cleargrid grdQUERY
                End If
            End If
        End If
        If Opt4.Value = True Then
            If chkWStock.Value = 1 Then
                Set rsPartmasDNPP = New ADODB.Recordset
                rsPartmasDNPP.Open "Select PMIS_PARTMAS.ONHAND,PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.DNPP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.DNPP,2) < ROUND(PMIS_PARTMAS.DNP,2) AND PMIS_PARTMAS.ONHAND > 0 order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                initPARTSINQUIRYGrid
                If Not rsPartmasDNPP.EOF And Not rsPartmasDNPP.BOF Then
                    KCNT = 0
                    Screen.MousePointer = 11
                    rsPartmasDNPP.MoveFirst
                    cmdSearchPartNo.Enabled = False
                    cmdExit.Enabled = False
                    Do While Not rsPartmasDNPP.EOF
                        KCNT = KCNT + 1
                        grdQUERY.AddItem Null2String(rsPartmasDNPP!PARTNO) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTDESC) & Chr(9) & _
                                         ToDoubleNumber(Null2String(rsPartmasDNPP!DEALERDNP)) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTNUMBER) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!DESCRIPTIO) & Chr(9) & _
                                         ToDoubleNumber(N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP))
                        rsPartmasDNPP.MoveNext
                        If KCNT = 1 Then grdQUERY.RemoveItem 1
                        progCPB.Value = (KCNT / rsPartmasDNPP.RecordCount) * 100
                        labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
                        DoEvents
                        DoEvents
                    Loop
                    LogAudit "V", "Parts DNP Comparision"
                    cmdSearchPartNo.Enabled = True
                    cmdExit.Enabled = True
                    Screen.MousePointer = 0
                Else
                    cleargrid grdQUERY
                End If
            Else
                Set rsPartmasDNPP = New ADODB.Recordset
                rsPartmasDNPP.Open "Select PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.DNPP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.DNPP,2) < ROUND(PMIS_PARTMAS.DNP,2) order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                initPARTSINQUIRYGrid
                If Not rsPartmasDNPP.EOF And Not rsPartmasDNPP.BOF Then
                    KCNT = 0
                    Screen.MousePointer = 11
                    rsPartmasDNPP.MoveFirst
                    cmdSearchPartNo.Enabled = False
                    cmdExit.Enabled = False
                    Do While Not rsPartmasDNPP.EOF
                        KCNT = KCNT + 1
                        grdQUERY.AddItem Null2String(rsPartmasDNPP!PARTNO) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTDESC) & Chr(9) & _
                                         ToDoubleNumber(Null2String(rsPartmasDNPP!DEALERDNP)) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!PARTNUMBER) & Chr(9) & _
                                         Null2String(rsPartmasDNPP!DESCRIPTIO) & Chr(9) & _
                                         ToDoubleNumber(N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP))
                        rsPartmasDNPP.MoveNext
                        If KCNT = 1 Then grdQUERY.RemoveItem 1
                        progCPB.Value = (KCNT / rsPartmasDNPP.RecordCount) * 100
                        labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
                        DoEvents
                        DoEvents
                    Loop
                    LogAudit "V", "Parts DNP Comparision"
                    cmdSearchPartNo.Enabled = True
                    cmdExit.Enabled = True
                    Screen.MousePointer = 0
                Else
                    cleargrid grdQUERY
                End If
            End If
        End If
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdUpdate_Click()
    Set rsPartmasDNPP = New ADODB.Recordset
    rsPartmasDNPP.Open "Select PMIS_PARTMAS.ONHAND,PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.SRP,2) as DEALERDNP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.SRP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(ROUND(PMIS_DNPP.SRP,2),2) > round(ROUND(PMIS_PARTMAS.SRP,2),2) AND PMIS_PARTMAS.ONHAND > 0 order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartmasDNPP.EOF And Not rsPartmasDNPP.BOF Then
        Screen.MousePointer = 11
        rsPartmasDNPP.MoveFirst
        cmdSearchPartNo.Enabled = False
        cmdExit.Enabled = False
        KCNT = 0
        Do While Not rsPartmasDNPP.EOF
            KCNT = KCNT + 1
            gconDMIS.Execute "update PMIS_PARTMAS set DNP = " & N2Str2Zero(rsPartmasDNPP!DISTRIBUTORDNP) & _
                           " where PARTNO = " & N2Str2Null(rsPartmasDNPP!PARTNO)

            rsPartmasDNPP.MoveNext
            progCPB.Value = (KCNT / rsPartmasDNPP.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
            DoEvents
        Loop
        NEW_LogAudit "R", "DEALER DISTRIBUTOR DNP COMPARISON", "", "", "", "", "", ""
        cmdSearchPartNo.Enabled = True
        cmdExit.Enabled = True
        Screen.MousePointer = 0
    End If
    initPARTSINQUIRYGrid
    FillPARTSINQUIRYGrid
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    cboRank.Clear
    cboRank.AddItem "ALL"
    cboRank.AddItem "A1"
    cboRank.AddItem "A2"
    cboRank.AddItem "A3"
    cboRank.AddItem "B"
    cboRank.AddItem "C"
    cboRank.AddItem "D"
    cboRank.AddItem "E1"
    cboRank.AddItem "E2"
    cboRank.AddItem "E3"
    cboRank.AddItem "E4"
    cboRank.AddItem "E5"
    cboRank.AddItem "F"
    cboRank.Text = "ALL"
    opt2.Value = True
    chkWStock.Value = 1
    If LOGLEVEL = "ADM" Then
        cmdUpdate.Enabled = False
    Else
        cmdUpdate.Enabled = False
    End If
    Set rsPartmasDNPP = New ADODB.Recordset
    Set rsPartmasDNPP = New ADODB.Recordset
    rsPartmasDNPP.Open "Select PMIS_PARTMAS.ONHAND,PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC, " & _
                     " ROUND(PMIS_PARTMAS.DNP,2) as DEALERDNP , " & _
                     " PMIS_DNPP.PARTNUMBER , PMIS_DNPP.DESCRIPTIO, ROUND(PMIS_DNPP.DNPP,2) as DISTRIBUTORDNP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.DNPP,2) > ROUND((PMIS_PARTMAS.MAC * " & ConvertToBIRDecimalFormat(VAT_RATE) & "),2) AND PMIS_PARTMAS.ONHAND > 0 order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    initPARTSINQUIRYGrid
    FillPARTSINQUIRYGrid
    Screen.MousePointer = 0
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearchPARTNO_Click()
    Dim FOUNDTEXT                                                     As String
    Dim FOUNDNUM                                                      As Long
    Dim findStr                                                       As String
    FOUNDNUM = 0
    grdQUERY.Col = 0
    findStr = InputSpeechBox("Please Input Part Number to Search", grdQUERY.Text)
    If findStr <> "" Then
        If Not PARTSINQUIRYBFound(findStr) Then
            MsgSpeechBox "Part Number " & findStr & " Not Found!"
        Else
            FOUNDNUM = rsPartmasDNPP.AbsolutePosition
            grdQUERY.Row = FOUNDNUM
            grdQUERY.RowSel = FOUNDNUM
            On Error Resume Next
            grdQUERY.SetFocus
            FOUNDTEXT = grdQUERY.Text
            grdQUERY.TopRow = FOUNDNUM
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub


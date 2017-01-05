VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmPMISInquiry_PartsSRPComparison 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PARTS SRP COMPARISON"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   ForeColor       =   &H00DEDFDE&
   Icon            =   "PartsSRPComparison.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   11640
   Begin wizButton.cmd cmdQuery 
      Height          =   525
      Left            =   8880
      TabIndex        =   6
      ToolTipText     =   "Process specified query"
      Top             =   2250
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
      MICON           =   "PartsSRPComparison.frx":030A
   End
   Begin Crystal.CrystalReport rptSRPQuery 
      Left            =   11100
      Top             =   4710
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
      TabIndex        =   5
      Top             =   1920
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
      TabIndex        =   14
      ToolTipText     =   "Select one option from list"
      Top             =   300
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
         Left            =   60
         TabIndex        =   1
         Top             =   270
         Value           =   -1  'True
         Width           =   2595
      End
      Begin VB.OptionButton Opt4 
         Caption         =   "DIST. SRP < DEALER SRP"
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
         Left            =   60
         TabIndex        =   4
         Top             =   1140
         Width           =   2595
      End
      Begin VB.OptionButton Opt3 
         Caption         =   "DIST. SRP = DEALER SRP"
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
         Left            =   60
         TabIndex        =   3
         Top             =   840
         Width           =   2595
      End
      Begin VB.OptionButton Opt2 
         Caption         =   "DIST. SRP > DEALER SRP"
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
         Left            =   60
         TabIndex        =   2
         Top             =   540
         Width           =   2595
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdQUERY 
      Height          =   5625
      Left            =   60
      TabIndex        =   0
      Top             =   510
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9922
      _Version        =   393216
      Cols            =   5
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
      TabIndex        =   7
      ToolTipText     =   "Search for a part number"
      Top             =   2850
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
      MICON           =   "PartsSRPComparison.frx":0624
   End
   Begin wizButton.cmd cmdPrintQuery 
      Height          =   525
      Left            =   8880
      TabIndex        =   8
      ToolTipText     =   "Print current query"
      Top             =   3450
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
      MICON           =   "PartsSRPComparison.frx":093E
   End
   Begin wizButton.cmd cmdUpdate 
      Height          =   525
      Left            =   8880
      TabIndex        =   9
      ToolTipText     =   "Update Dealer Master File"
      Top             =   4050
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   926
      TX              =   "Update Dealer Master File"
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
      MICON           =   "PartsSRPComparison.frx":0C58
   End
   Begin wizButton.cmd cmdExit 
      Height          =   525
      Left            =   8880
      TabIndex        =   10
      ToolTipText     =   "Close window"
      Top             =   4650
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
      MICON           =   "PartsSRPComparison.frx":0F72
   End
   Begin wizProgBar.Prg progCPB 
      Height          =   345
      Left            =   9030
      TabIndex        =   16
      ToolTipText     =   "Process progress"
      Top             =   5700
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   609
      Picture         =   "PartsSRPComparison.frx":128C
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "PartsSRPComparison.frx":12A8
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
   Begin VB.Label labCPB 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   8910
      TabIndex        =   15
      Top             =   5220
      Width           =   2625
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "DISTRIBUTOR PARTS MASTERFILE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4260
      TabIndex        =   13
      Top             =   90
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
      Height          =   345
      Left            =   90
      TabIndex        =   12
      Top             =   90
      Width           =   4125
   End
   Begin VB.Label labAydi 
      Caption         =   "Label1"
      Height          =   195
      Left            =   7440
      TabIndex        =   11
      Top             =   4860
      Visible         =   0   'False
      Width           =   30
   End
End
Attribute VB_Name = "frmPMISInquiry_PartsSRPComparison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPartmasPMIS_DNPP                                 As ADODB.Recordset
Dim KCNT                                               As Double

Function PARTSINQUIRYBFound(ByVal str2find) As Boolean
    On Error GoTo BFoundErr
    Dim result                                         As Boolean
    Dim rsBClone                                       As ADODB.Recordset
    result = False
    If Not IsNull(str2find) Then
        Set rsBClone = New ADODB.Recordset
        Set rsBClone = rsPartmasPMIS_DNPP.Clone

        rsBClone.Find "PARTNO = '" & str2find & "'"
        result = Not rsBClone.EOF
        If result Then
            rsPartmasPMIS_DNPP.Bookmark = rsBClone.Bookmark
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
        .FormatString = "Part Number  | Part Desc               |    SRP     | " & _
                        "Part Number  | Part Desc               |    SRP     "
    End With
End Sub

Sub FillPARTSINQUIRYGrid()
    On Error GoTo ErrorCode
    KCNT = "0"
    If Not rsPartmasPMIS_DNPP.EOF And Not rsPartmasPMIS_DNPP.BOF Then
        Screen.MousePointer = 11
        rsPartmasPMIS_DNPP.MoveFirst
        cmdSearchPartNo.Enabled = False
        cmdExit.Enabled = False
        Do While Not rsPartmasPMIS_DNPP.EOF
            KCNT = KCNT + 1
            grdQUERY.AddItem Null2String(rsPartmasPMIS_DNPP!PARTNO) & Chr(9) & _
                             Null2String(rsPartmasPMIS_DNPP!PARTDESC) & Chr(9) & _
                             ToDoubleNumber(N2Str2Zero(rsPartmasPMIS_DNPP!DEALERSRP)) & Chr(9) & _
                             Null2String(rsPartmasPMIS_DNPP!PARTNUMBER) & Chr(9) & _
                             Null2String(rsPartmasPMIS_DNPP!DESCRIPTIO) & Chr(9) & _
                             ToDoubleNumber(N2Str2Zero(rsPartmasPMIS_DNPP!DISTRIBUTORSRP))
            rsPartmasPMIS_DNPP.MoveNext
            If KCNT = 1 Then grdQUERY.RemoveItem 1
            progCPB.Value = (KCNT / rsPartmasPMIS_DNPP.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
            DoEvents
            DoEvents
        Loop
        LogAudit "V", "Parts SRP Comparision"
        cmdSearchPartNo.Enabled = True
        cmdExit.Enabled = True
        Screen.MousePointer = 0
    Else
        cleargrid grdQUERY
    End If
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdPrintQuery_Click()
    If Opt1.Value = True Then
        If chkWStock.Value = 1 Then
            rptSRPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptSRPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptSRPQuery, PMIS_REPORT_PATH & "SRPcomparison.rpt", "{PMIS_PARTMAS.ONHAND} > 0", DMIS_REPORT_Connection, 1
        Else
            rptSRPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptSRPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptSRPQuery, PMIS_REPORT_PATH & "SRPcomparison.rpt", "", DMIS_REPORT_Connection, 1
        End If
    End If
    If Opt2.Value = True Then
        If chkWStock.Value = 1 Then
            rptSRPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptSRPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptSRPQuery, PMIS_REPORT_PATH & "SRPcomparison.rpt", "ROUND({PMIS_DNPP.SRP},2) > ROUND({PMIS_PARTMAS.SRP},2) AND {PMIS_PARTMAS.ONHAND} > 0", DMIS_REPORT_Connection, 1
        Else
            rptSRPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptSRPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptSRPQuery, PMIS_REPORT_PATH & "SRPcomparison.rpt", "ROUND({PMIS_DNPP.SRP},2) > ROUND({PMIS_PARTMAS.SRP},2)", DMIS_REPORT_Connection, 1
        End If
    End If
    If Opt3.Value = True Then
        If chkWStock.Value = 1 Then
            rptSRPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptSRPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptSRPQuery, PMIS_REPORT_PATH & "SRPcomparison.rpt", "ROUND({PMIS_DNPP.SRP},2) = ROUND({PMIS_PARTMAS.SRP},2) AND {PMIS_PARTMAS.ONHAND} > 0", DMIS_REPORT_Connection, 1
        Else
            rptSRPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptSRPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptSRPQuery, PMIS_REPORT_PATH & "SRPcomparison.rpt", "ROUND({PMIS_DNPP.SRP},2) = ROUND({PMIS_PARTMAS.SRP},2)", DMIS_REPORT_Connection, 1
        End If
    End If
    If Opt4.Value = True Then
        If chkWStock.Value = 1 Then
            rptSRPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptSRPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptSRPQuery, PMIS_REPORT_PATH & "SRPcomparison.rpt", "ROUND({PMIS_DNPP.SRP},2) < ROUND({PMIS_PARTMAS.SRP},2) AND {PMIS_PARTMAS.ONHAND} > 0", DMIS_REPORT_Connection, 1
        Else
            rptSRPQuery.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptSRPQuery.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptSRPQuery, PMIS_REPORT_PATH & "SRPcomparison.rpt", "ROUND({PMIS_DNPP.SRP},2) < ROUND({PMIS_PARTMAS.SRP},2)", DMIS_REPORT_Connection, 1
        End If
    End If
End Sub

Private Sub cmdQuery_Click()
    Screen.MousePointer = 11
    cleargrid grdQUERY
    If Opt1.Value = True Then
        If chkWStock.Value = 1 Then
            Set rsPartmasPMIS_DNPP = New ADODB.Recordset
            rsPartmasPMIS_DNPP.Open "Select PMIS_PARTMAS.ONHAND,PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.SRP,2) as DEALERSRP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.SRP,2) as DISTRIBUTORSRP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where PMIS_PARTMAS.ONHAND > 0 order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            initPARTSINQUIRYGrid
            FillPARTSINQUIRYGrid
        Else
            Set rsPartmasPMIS_DNPP = New ADODB.Recordset
            rsPartmasPMIS_DNPP.Open "Select PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.SRP,2) as DEALERSRP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.SRP,2) as DISTRIBUTORSRP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            initPARTSINQUIRYGrid
            FillPARTSINQUIRYGrid
        End If
    End If
    If Opt2.Value = True Then
        If chkWStock.Value = 1 Then
            Set rsPartmasPMIS_DNPP = New ADODB.Recordset
            rsPartmasPMIS_DNPP.Open "Select PMIS_PARTMAS.ONHAND,PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.SRP,2) as DEALERSRP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.SRP,2) as DISTRIBUTORSRP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.SRP,2) > ROUND(PMIS_PARTMAS.SRP,2) AND PMIS_PARTMAS.ONHAND > 0 order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            initPARTSINQUIRYGrid
            FillPARTSINQUIRYGrid
        Else
            Set rsPartmasPMIS_DNPP = New ADODB.Recordset
            rsPartmasPMIS_DNPP.Open "Select PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.SRP,2) as DEALERSRP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.SRP,2) as DISTRIBUTORSRP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.SRP,2) > ROUND(PMIS_PARTMAS.SRP,2) order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            initPARTSINQUIRYGrid
            FillPARTSINQUIRYGrid
        End If
    End If
    If Opt3.Value = True Then
        If chkWStock.Value = 1 Then
            Set rsPartmasPMIS_DNPP = New ADODB.Recordset
            rsPartmasPMIS_DNPP.Open "Select PMIS_PARTMAS.ONHAND,PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.SRP,2) as DEALERSRP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.SRP,2) as DISTRIBUTORSRP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.SRP,2) = ROUND(PMIS_PARTMAS.SRP,0) AND PMIS_PARTMAS.ONHAND > 0 order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            initPARTSINQUIRYGrid
            FillPARTSINQUIRYGrid
        Else
            Set rsPartmasPMIS_DNPP = New ADODB.Recordset
            rsPartmasPMIS_DNPP.Open "Select PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.SRP,2) as DEALERSRP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.SRP,2) as DISTRIBUTORSRP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER  where ROUND(PMIS_DNPP.SRP,2) = ROUND(PMIS_PARTMAS.SRP,0) order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            initPARTSINQUIRYGrid
            FillPARTSINQUIRYGrid
        End If
    End If
    If Opt4.Value = True Then
        If chkWStock.Value = 1 Then
            Set rsPartmasPMIS_DNPP = New ADODB.Recordset
            rsPartmasPMIS_DNPP.Open "Select PMIS_PARTMAS.ONHAND,PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.SRP,2) as DEALERSRP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.SRP,2) as DISTRIBUTORSRP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.SRP,2) < ROUND(PMIS_PARTMAS.SRP,2) AND PMIS_PARTMAS.ONHAND > 0 order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            initPARTSINQUIRYGrid
            FillPARTSINQUIRYGrid
        Else
            Set rsPartmasPMIS_DNPP = New ADODB.Recordset
            rsPartmasPMIS_DNPP.Open "Select PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.SRP,2) as DEALERSRP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.SRP,2) as DISTRIBUTORSRP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER  where ROUND(PMIS_DNPP.SRP,2) < ROUND(PMIS_PARTMAS.SRP,2) order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            initPARTSINQUIRYGrid
            FillPARTSINQUIRYGrid
        End If
    End If
    Screen.MousePointer = 0
End Sub

Private Sub cmdUpdate_Click()
    Set rsPartmasPMIS_DNPP = New ADODB.Recordset
    rsPartmasPMIS_DNPP.Open "Select PMIS_PARTMAS.ONHAND,PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.SRP,2) as DEALERSRP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO,ROUND(PMIS_DNPP.SRP,2) as DISTRIBUTORSRP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.SRP,2) > ROUND(PMIS_PARTMAS.SRP,2) AND PMIS_PARTMAS.ONHAND > 0 order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartmasPMIS_DNPP.EOF And Not rsPartmasPMIS_DNPP.BOF Then
        Screen.MousePointer = 11
        rsPartmasPMIS_DNPP.MoveFirst
        cmdSearchPartNo.Enabled = False
        cmdExit.Enabled = False
        KCNT = 0
        Do While Not rsPartmasPMIS_DNPP.EOF
            KCNT = KCNT + 1
            gconDMIS.Execute "update PMIS_PARTMAS set srp = " & N2Str2Zero(rsPartmasPMIS_DNPP!DISTRIBUTORSRP) & _
                           " where PARTNO = " & N2Str2Null(rsPartmasPMIS_DNPP!PARTNO)
            rsPartmasPMIS_DNPP.MoveNext
            progCPB.Value = (KCNT / rsPartmasPMIS_DNPP.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed [Total Updated Record(s) = " & KCNT
            DoEvents
        Loop
        NEW_LogAudit "R", "DEALER DISTRIBUTOR SRP COMPARISON", "", "", "", "", "", ""
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
    Opt2.Value = True
    chkWStock.Value = 1
    If LOGLEVEL = "ADM" Then
        cmdUpdate.Enabled = True
    Else
        cmdUpdate.Enabled = False
    End If
    Set rsPartmasPMIS_DNPP = New ADODB.Recordset
    rsPartmasPMIS_DNPP.Open "Select PMIS_PARTMAS.ONHAND,PMIS_PARTMAS.PARTNO,PMIS_PARTMAS.PARTDESC,ROUND(PMIS_PARTMAS.SRP,0) as DEALERSRP,PMIS_DNPP.PARTNUMBER,PMIS_DNPP.DESCRIPTIO, ROUND(PMIS_DNPP.SRP,2) as DISTRIBUTORSRP from PMIS_PARTMAS inner join PMIS_DNPP on PMIS_PARTMAS.PARTNO = PMIS_DNPP.PARTNUMBER where ROUND(PMIS_DNPP.SRP,2) > ROUND(PMIS_PARTMAS.SRP,2) AND PMIS_PARTMAS.ONHAND > 0 order by PMIS_PARTMAS.PARTNO asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    initPARTSINQUIRYGrid
    FillPARTSINQUIRYGrid
    Screen.MousePointer = 0
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearchPARTNO_Click()
    Dim FOUNDTEXT                                      As String
    Dim FOUNDNUM                                       As Long
    Dim findStr                                        As String
    FOUNDNUM = 0
    grdQUERY.Col = 0
    findStr = InputSpeechBox("Please Input Part Number to Search", grdQUERY.Text)
    If findStr <> "" Then
        If Not PARTSINQUIRYBFound(findStr) Then
            MsgSpeechBox "Part Number " & findStr & " Not Found!"
        Else
            FOUNDNUM = rsPartmasPMIS_DNPP.AbsolutePosition
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


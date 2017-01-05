VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCSMSTechnicianPerformanceReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tech Performance Report"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3825
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMSTechnicianPerformanceReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2460
   ScaleWidth      =   3825
   Begin VB.ComboBox cboTechnicianPerformance_Report 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   420
      Width           =   3435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   825
      Left            =   2910
      MouseIcon       =   "frmCSMSTechnicianPerformanceReport.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSTechnicianPerformanceReport.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Close Window"
      Top             =   1530
      Width           =   705
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Contractor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      Top             =   90
      Width           =   1425
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Technician"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   90
      Value           =   -1  'True
      Width           =   1425
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4770
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Select month from the list"
      Top             =   930
      Width           =   1965
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4770
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Select year from the list"
      Top             =   1350
      Width           =   1965
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7350
      MouseIcon       =   "frmCSMSTechnicianPerformanceReport.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSTechnicianPerformanceReport.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   1860
      Width           =   795
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6570
      MouseIcon       =   "frmCSMSTechnicianPerformanceReport.frx":197C
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSTechnicianPerformanceReport.frx":1ACE
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print Report"
      Top             =   1860
      Width           =   795
   End
   Begin MSComCtl2.DTPicker dtpFROM 
      Height          =   375
      Left            =   180
      TabIndex        =   9
      Top             =   1020
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   97583105
      CurrentDate     =   39646
   End
   Begin Crystal.CrystalReport rptTechnician_Performance_Report 
      Left            =   180
      Top             =   1740
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Technician Performance Report"
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker dtpTO 
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   1020
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   97583105
      CurrentDate     =   39646
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   825
      Left            =   2220
      MouseIcon       =   "frmCSMSTechnicianPerformanceReport.frx":1F6D
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMSTechnicianPerformanceReport.frx":20BF
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Print Report"
      Top             =   1530
      Width           =   705
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   210
      TabIndex        =   14
      Top             =   780
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   2
      Left            =   1965
      TabIndex        =   13
      Top             =   780
      Width           =   210
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   4020
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   4050
      TabIndex        =   4
      Top             =   1380
      Width           =   735
   End
End
Attribute VB_Name = "frmCSMSTechnicianPerformanceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Sub FillCombo()
'    Dim tmp_value                                                     As String
'    Dim rsTechnician_Performance_Report                               As ADODB.Recordset
'    tmp_value = ""
'    Set rsTechnician_Performance_Report = New ADODB.Recordset
'    rsTechnician_Performance_Report.Open "Select Tech_Name from CSMS_JobClock order by Tech_Name asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
'    If Not rsTechnician_Performance_Report.EOF And Not rsTechnician_Performance_Report.BOF Then
'        rsTechnician_Performance_Report.MoveFirst
'        cboTechnicianPerformance_Report.Clear
'        cboTechnicianPerformance_Report.AddItem "All"
'        Do While Not rsTechnician_Performance_Report.EOF
'            If tmp_value = Null2String(rsTechnician_Performance_Report!Tech_Name) Then
'                rsTechnician_Performance_Report.MoveNext
'            Else
'                cboTechnicianPerformance_Report.AddItem Null2String(rsTechnician_Performance_Report!Tech_Name)
'                tmp_value = Null2String(rsTechnician_Performance_Report!Tech_Name)
'                rsTechnician_Performance_Report.MoveNext
'            End If
'        Loop
'    End If
'    Set rsTechnician_Performance_Report = Nothing
'End Sub

Function SetEmpNO(XXX As String) As String
    Dim rsTechnician                                   As New ADODB.Recordset
    Set rsTechnician = gconDMIS.Execute("Select * from CSMS_vw_Technician where tech_name = '" & XXX & "'")
    If Not rsTechnician.EOF And Not rsTechnician.BOF Then
        SetEmpNO = LTrim(RTrim(Null2String(rsTechnician!EMPNO)))
    End If
End Function

Function SetCodeNO(XXX As String) As String
    Dim rsTechnician                                   As New ADODB.Recordset
    Set rsTechnician = gconDMIS.Execute("Select * from CSMS_Contractor where COMPANYname = '" & XXX & "'")
    If Not rsTechnician.EOF And Not rsTechnician.BOF Then
        SetCodeNO = LTrim(RTrim(Null2String(rsTechnician!Code)))
    End If
End Function

Function SetVendorCode(XXX As String) As String
    Dim rsTechnician                                   As New ADODB.Recordset
    Set rsTechnician = gconDMIS.Execute("Select CODE from ALL_VENDOR_TABLE where NAMEOFVENDOR = " & N2Str2Null(XXX) & "")
    If Not rsTechnician.EOF And Not rsTechnician.BOF Then
        SetVendorCode = LTrim(RTrim(Null2String(rsTechnician!Code)))
    End If
End Function

Sub cmdPrint_Click()
    On Error GoTo ErrorCode

    Dim rsTechnician_Performance_Report                As New ADODB.Recordset
    Set rsTechnician_Performance_Report = gconDMIS.Execute("Select * from CSMS_vw_RO_Det_For_TechPerformance")

    If cboTechnicianPerformance_Report.Text = "ALL" Then
        If Not rsTechnician_Performance_Report.EOF And Not rsTechnician_Performance_Report.BOF Then
            Screen.MousePointer = 11

            'JUN 02/05/2005
            rptTechnician_Performance_Report.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptTechnician_Performance_Report.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptTechnician_Performance_Report.Formulas(2) = "Printedby = '" & LOGNAME & "'"
            
            PrintSQLReport rptTechnician_Performance_Report, CSMS_REPORT_PATH & "TechnicianPerformance.rpt", "month({csms_repor.dte_comp}) = " & What_month(cboMonth) & " and year({csms_repor.dte_comp}) = " & cboYear & "", CSMS_REPORT_CONNECTION, 1

            LogAudit "V", "TECHNICIAN PERFORMANCE - REPORT", cboTechnicianPerformance_Report
            Screen.MousePointer = 0
        Else
            ShowNoRecord
            On Error Resume Next
            cboTechnicianPerformance_Report.SetFocus
            Exit Sub
        End If
    Else
        If Not rsTechnician_Performance_Report.EOF And Not rsTechnician_Performance_Report.BOF Then
            Screen.MousePointer = 11

            'JUN 02/05/2005
            rptTechnician_Performance_Report.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptTechnician_Performance_Report.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptTechnician_Performance_Report.Formulas(2) = "Printedby = '" & LOGNAME & "'"


            PrintSQLReport rptTechnician_Performance_Report, CSMS_REPORT_PATH & "TechnicianPerformance.rpt", "month({csms_repor.dte_comp}) = " & What_month(cboMonth) & " and year({csms_repor.dte_comp}) = " & cboYear & " and {CSMS_vw_RO_Det_For_TechPerformance.TECHNICIAN} = '" & SetEmpNO(cboTechnicianPerformance_Report) & "' and {csms}", CSMS_REPORT_CONNECTION, 1
            LogAudit "G", "TECHNICIAN PERFORMANCE - REPORT"

            Screen.MousePointer = 0
        Else
            ShowNoRecord
            cboTechnicianPerformance_Report.SetFocus
            Exit Sub
        End If
    End If
    Exit Sub
ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Sub FillCombo()
    Dim rstmp                                          As New ADODB.Recordset
    Dim NEYM                                           As String

    cboTechnicianPerformance_Report.Clear
    cboTechnicianPerformance_Report.AddItem "ALL"
    Set rstmp = gconDMIS.Execute("SELECT * FROM CSMS_vw_Technician order by tech_name")
    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            'NEYM = Null2String(RSTMP!lastname) & ", " & Null2String(RSTMP!Firstname) & " " & Left(Null2String(RSTMP!MIDDLENAME), 1) & "."

            cboTechnicianPerformance_Report.AddItem Null2String(rstmp!TECH_NAME)
            rstmp.MoveNext
        Loop
    End If
    cboTechnicianPerformance_Report.Text = "ALL"
    Set rstmp = Nothing
End Sub

Sub FillComboContractor()
    Dim rstmp                                          As New ADODB.Recordset
    Dim NEYM                                           As String

    cboTechnicianPerformance_Report.Clear
    cboTechnicianPerformance_Report.AddItem "ALL"
    'Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_CONTRACTOR order by COMPANYNAME")
    Set rstmp = gconDMIS.Execute("SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE order by NAMEOFVENDOR")
    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            cboTechnicianPerformance_Report.AddItem Null2String(rstmp!nameofvendor)
            rstmp.MoveNext
        Loop
    End If

    cboTechnicianPerformance_Report.Text = "ALL"
    Set rstmp = Nothing
End Sub

Private Sub cboTechnicianPerformance_Report_KeyPress(KeyAscii As Integer)
    'Call cmdPrint_Click
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    On Error GoTo ErrorCode

    Screen.MousePointer = 11

    If dtpTO.Value < dtpFROM.Value Then
        MsgBox "Invalid Range Format", vbInformation, "CSMS"
        dtpFROM.SetFocus
        Exit Sub
    End If
    Dim VRANGE                                         As String
    Dim rsTechnician_Performance_Report                As New ADODB.Recordset
    
    Set rsTechnician_Performance_Report = gconDMIS.Execute("Select * from CSMS_REPOR WHERE DTE_COMP BETWEEN '" & dtpFROM.Value & "' and '" & dtpTO.Value & "'")
    VRANGE = "From " & dtpFROM.Value & " To " & dtpTO.Value
    
    If Option1.Value = True Then                      'TECHNICIAN
        If Not (rsTechnician_Performance_Report.BOF And rsTechnician_Performance_Report.EOF) Then
            If cboTechnicianPerformance_Report.Text = "ALL" Then
                'JUN 02/05/2005
                rptTechnician_Performance_Report.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptTechnician_Performance_Report.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptTechnician_Performance_Report.Formulas(2) = "Printedby = '" & LOGNAME & "'"
                rptTechnician_Performance_Report.Formulas(3) = "RANGE = '" & VRANGE & "' "
                rptTechnician_Performance_Report.WindowTitle = "TECHNICIAN PERFORMANCE REPORT"
                
                If COMPANY_CODE = "MGS" Then
                    PrintSQLReport rptTechnician_Performance_Report, CSMS_REPORT_PATH & "TechnicianPerformance.rpt", "{CSMS_VW_LABOR_WARRANTY_COMPANY.DATE} >= date(" & Year(dtpFROM.Value) & "," & Month(dtpFROM.Value) & "," & Day(dtpFROM.Value) & ") AND {CSMS_VW_LABOR_WARRANTY_COMPANY.DATE} <= DATE(" & Year(dtpTO.Value) & "," & Month(dtpTO.Value) & "," & Day(dtpTO.Value) & ")", CSMS_REPORT_CONNECTION, 1
                Else
                    PrintSQLReport rptTechnician_Performance_Report, CSMS_REPORT_PATH & "TechnicianPerformance.rpt", "{csms_repor.dte_comp} >= date(" & Year(dtpFROM.Value) & "," & Month(dtpFROM.Value) & "," & Day(dtpFROM.Value) & ") AND {CSMS_REPOR.DTE_COMP} <= DATE(" & Year(dtpTO.Value) & "," & Month(dtpTO.Value) & "," & Day(dtpTO.Value) & ")", CSMS_REPORT_CONNECTION, 1
                End If
                
                If COMPANY_CODE = "HPI" Then
                    PrintSQLReport rptTechnician_Performance_Report, CSMS_REPORT_PATH & "ContractorPerformance_JobCost2.rpt", "{csms_repor.dte_comp} >= date(" & Year(dtpFROM.Value) & "," & Month(dtpFROM.Value) & "," & Day(dtpFROM.Value) & ") AND {CSMS_REPOR.DTE_COMP} <= DATE(" & Year(dtpTO.Value) & "," & Month(dtpTO.Value) & "," & Day(dtpTO.Value) & ")", CSMS_REPORT_CONNECTION, 1
                End If
            Else
                'JUN 02/05/2005
                rptTechnician_Performance_Report.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptTechnician_Performance_Report.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptTechnician_Performance_Report.Formulas(2) = "Printedby = '" & LOGNAME & "'"
                rptTechnician_Performance_Report.Formulas(3) = "RANGE = '" & VRANGE & "' "
                rptTechnician_Performance_Report.WindowTitle = "TECHNICIAN PERFORMANCE REPORT"
                
                If COMPANY_CODE = "MGS" Then
                    PrintSQLReport rptTechnician_Performance_Report, CSMS_REPORT_PATH & "TechnicianPerformance.rpt", "{CSMS_VW_LABOR_WARRANTY_COMPANY.DATE} >= date(" & Year(dtpFROM.Value) & "," & Month(dtpFROM.Value) & "," & Day(dtpFROM.Value) & ") AND {CSMS_VW_LABOR_WARRANTY_COMPANY.DATE} <= DATE(" & Year(dtpTO.Value) & "," & Month(dtpTO.Value) & "," & Day(dtpTO.Value) & ") AND  {CSMS_VW_LABOR_WARRANTY_COMPANY.EMPNO} = '" & SetEmpNO(cboTechnicianPerformance_Report) & "'", CSMS_REPORT_CONNECTION, 1
                Else
                    PrintSQLReport rptTechnician_Performance_Report, CSMS_REPORT_PATH & "TechnicianPerformance.rpt", "{csms_repor.dte_comp} >= date(" & Year(dtpFROM.Value) & "," & Month(dtpFROM.Value) & "," & Day(dtpFROM.Value) & ") AND {CSMS_REPOR.DTE_COMP} <= DATE(" & Year(dtpTO.Value) & "," & Month(dtpTO.Value) & "," & Day(dtpTO.Value) & ") AND  {CSMS_vw_Technician.empno} = '" & SetEmpNO(cboTechnicianPerformance_Report) & "'", CSMS_REPORT_CONNECTION, 1
                End If
                
                If COMPANY_CODE = "HPI" Then
                    PrintSQLReport rptTechnician_Performance_Report, CSMS_REPORT_PATH & "ContractorPerformance_JobCost.rpt", "{csms_repor.dte_comp} >= date(" & Year(dtpFROM.Value) & "," & Month(dtpFROM.Value) & "," & Day(dtpFROM.Value) & ") AND {CSMS_REPOR.DTE_COMP} <= DATE(" & Year(dtpTO.Value) & "," & Month(dtpTO.Value) & "," & Day(dtpTO.Value) & ") AND  {CSMS_CONTRACTOR.CODE} = '" & SetCodeNO(cboTechnicianPerformance_Report) & "'", CSMS_REPORT_CONNECTION, 1
                End If
            End If
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "TECHNICIAN REPORT", "", "", "", "TECHNICIAN PERFORMANCE : " & cboTechnicianPerformance_Report & " : " & dtpFROM & " - " & dtpTO, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        Else
            Call ShowNoRecord
            On Error Resume Next
            cboTechnicianPerformance_Report.SetFocus
        End If
    Else                                              'CONTRACTOR
        If Not (rsTechnician_Performance_Report.BOF And rsTechnician_Performance_Report.EOF) Then
            If cboTechnicianPerformance_Report.Text = "ALL" Then
                'JUN 02/05/2005
                rptTechnician_Performance_Report.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptTechnician_Performance_Report.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptTechnician_Performance_Report.Formulas(2) = "Printedby = '" & LOGNAME & "'"
                rptTechnician_Performance_Report.Formulas(3) = "RANGE = '" & VRANGE & "' "
                rptTechnician_Performance_Report.WindowTitle = "CONTRACTOR PERFORMANCE REPORT"

                PrintSQLReport rptTechnician_Performance_Report, CSMS_REPORT_PATH & "ContractorPerformance_vendor.rpt", "{csms_repor.dte_comp} >= date(" & Year(dtpFROM.Value) & "," & Month(dtpFROM.Value) & "," & Day(dtpFROM.Value) & ") AND {CSMS_REPOR.DTE_COMP} <= DATE(" & Year(dtpTO.Value) & "," & Month(dtpTO.Value) & "," & Day(dtpTO.Value) & ")", CSMS_REPORT_CONNECTION, 1
                                
                'UPDATED BY: JUN
                'DATE UPDATED: 04022009
                'DESCRIPTION: UPDATED TO HPI DUE TO MONITORING OF JON COST
                    If COMPANY_CODE = "HPI" Then
                        PrintSQLReport rptTechnician_Performance_Report, CSMS_REPORT_PATH & "ContractorPerformance_JobCost_All.rpt", "{csms_repor.dte_comp} >= date(" & Year(dtpFROM.Value) & "," & Month(dtpFROM.Value) & "," & Day(dtpFROM.Value) & ") AND {CSMS_REPOR.DTE_COMP} <= DATE(" & Year(dtpTO.Value) & "," & Month(dtpTO.Value) & "," & Day(dtpTO.Value) & ")", CSMS_REPORT_CONNECTION, 1
                    End If
                'DATE UPDATED: 04022009
                'UPDATED BY: JUN
            Else
                'JUN 02/05/2005
                rptTechnician_Performance_Report.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptTechnician_Performance_Report.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptTechnician_Performance_Report.Formulas(2) = "Printedby = '" & LOGNAME & "'"
                rptTechnician_Performance_Report.Formulas(3) = "RANGE = '" & VRANGE & "' "
                rptTechnician_Performance_Report.WindowTitle = "CONTRACTOR PERFORMANCE REPORT"

                PrintSQLReport rptTechnician_Performance_Report, CSMS_REPORT_PATH & "ContractorPerformance_vendor.rpt", "{csms_repor.dte_comp} >= date(" & Year(dtpFROM.Value) & "," & Month(dtpFROM.Value) & "," & Day(dtpFROM.Value) & ") AND {CSMS_REPOR.DTE_COMP} <= DATE(" & Year(dtpTO.Value) & "," & Month(dtpTO.Value) & "," & Day(dtpTO.Value) & ") AND  {ALL_VENDOR_TABLE.CODE} = '" & SetVendorCode(cboTechnicianPerformance_Report) & "'", CSMS_REPORT_CONNECTION, 1
                
                'UPDATED BY: JUN
                'DATE UPDATED: 04022009
                'DESCRIPTION: UPDATED TO HPI DUE TO MONITORING OF JON COST
                    If COMPANY_CODE = "HPI" Then
                        PrintSQLReport rptTechnician_Performance_Report, CSMS_REPORT_PATH & "ContractorPerformance_JobCost_All.rpt", "{csms_repor.dte_comp} >= date(" & Year(dtpFROM.Value) & "," & Month(dtpFROM.Value) & "," & Day(dtpFROM.Value) & ") AND {CSMS_REPOR.DTE_COMP} <= DATE(" & Year(dtpTO.Value) & "," & Month(dtpTO.Value) & "," & Day(dtpTO.Value) & ") AND  {ALL_VENDOR_TABLE.CODE} = '" & SetVendorCode(cboTechnicianPerformance_Report) & "'", CSMS_REPORT_CONNECTION, 1
                    End If
                'DATE UPDATED: 04022009
                'UPDATED BY: JUN
            End If
            
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("V", "TECHNICIAN REPORT", "", "", "", "CONTRACTOR PERFORMANCE : " & cboTechnicianPerformance_Report & " : " & dtpFROM & " - " & dtpTO, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        Else
            Call ShowNoRecord
            On Error Resume Next
            cboTechnicianPerformance_Report.SetFocus
        End If
    End If

    Screen.MousePointer = 0
    Set rsTechnician_Performance_Report = Nothing
    
    Exit Sub
ErrorCode:
    ShowVBError
    Screen.MousePointer = 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
    
    Call fillcbomonth(cboMonth)
    Call FillCboMoreYear(cboYear)
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    Call FillCombo
    
    dtpFROM.Value = firstDay(Date)
    dtpTO.Value = Date
    
    Screen.MousePointer = 0
End Sub

Private Sub Option1_Click()
    On Error Resume Next
    If Option1.Value = True Then
        Call FillCombo
        cboTechnicianPerformance_Report.SetFocus
    End If
End Sub

Private Sub Option2_Click()
    'On Error Resume Next
    If Option2.Value = True Then
        Call FillComboContractor
        cboTechnicianPerformance_Report.SetFocus
    End If
End Sub


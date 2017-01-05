VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCSMSTechnician_Efficiency 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Technician Productivity"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3675
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Technician_Efficiency.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2460
   ScaleWidth      =   3675
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
      Left            =   120
      TabIndex        =   9
      Top             =   60
      Value           =   -1  'True
      Width           =   1425
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
      Left            =   2220
      TabIndex        =   8
      Top             =   60
      Width           =   1425
   End
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
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   390
      Width           =   3435
   End
   Begin MSComCtl2.DTPicker txtFrom 
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   1020
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   58785793
      CurrentDate     =   39546
   End
   Begin Crystal.CrystalReport rptTechnician 
      Left            =   150
      Top             =   1740
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Technician Efficiency Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   825
      Left            =   2820
      MouseIcon       =   "Technician_Efficiency.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "Technician_Efficiency.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Close Window"
      Top             =   1500
      Width           =   705
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   825
      Left            =   2130
      MouseIcon       =   "Technician_Efficiency.frx":13DF
      MousePointer    =   99  'Custom
      Picture         =   "Technician_Efficiency.frx":1531
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Report"
      Top             =   1500
      Width           =   705
   End
   Begin MSComCtl2.DTPicker txtTo 
      Height          =   345
      Left            =   1890
      TabIndex        =   6
      Top             =   1020
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   58785793
      CurrentDate     =   39546
   End
   Begin VB.Label Label1 
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
      Left            =   150
      TabIndex        =   2
      Top             =   810
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
      Left            =   1905
      TabIndex        =   1
      Top             =   810
      Width           =   210
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1650
      TabIndex        =   0
      Top             =   2970
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmCSMSTechnician_Efficiency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsRO_DET                                           As ADODB.Recordset
Dim rsREPOR                                            As ADODB.Recordset
Dim xlApp                                              As Excel.Application
Dim xlBook                                             As Excel.Workbook
Dim xlSheet                                            As Excel.Worksheet

Function ComputeAttendedHours(vEMPNO As String) As Double()
    Dim rsTMP                                          As New ADODB.Recordset
    Dim VTMP(1)                                        As Double
    Dim VTIME                                          As Double
    Dim VAVAI                                          As Double

    'Set rstmp = gconDMIS.Execute("SELECT * FROM HRMS_ATTEND WHERE EMPNO = '" & VEMPNO & "' AND MONTH(DATETODAY) = " & What_month(cboMonth) & " AND YEAR(DATETODAY) = " & cboYEAR & "")
    Set rsTMP = gconDMIS.Execute("SELECT * FROM HRMS_ATTEND WHERE EMPNO = '" & vEMPNO & "' AND DATETODAY BETWEEN '" & TXTFrom.Value & "' AND '" & txtTO.Value & "'")
    If Not (rsTMP.BOF And rsTMP.EOF) Then
        Do While Not rsTMP.EOF
            If Null2String(rsTMP!Shift) = "SHIFT1" Then
                If Not Null2String(rsTMP!INAM) = "" Then
                    If Not Null2String(rsTMP!OUTAM) = "" Then
                        VTIME = VTIME + DateDiff("N", rsTMP!INAM, rsTMP!OUTAM)
                    End If
                End If
            Else


            End If

            VAVAI = VAVAI + 7.5
            rsTMP.MoveNext
        Loop
    End If

    If Not VTIME < 0 Then
        VTMP(0) = VTIME / 60
        VTMP(1) = VAVAI
        ComputeAttendedHours = VTMP
    Else
        VTMP(0) = 0
        VTMP(1) = VAVAI
    End If

    ComputeAttendedHours = VTMP
    Set rsTMP = Nothing
End Function

Function SetEmpNO(XXX As String) As String
    Dim rsTechnician                                   As ADODB.Recordset
    Set rsTechnician = New ADODB.Recordset
    Set rsTechnician = gconDMIS.Execute("Select * from CSMS_vw_Technician where tech_name = '" & XXX & "'")
    If Not rsTechnician.EOF And Not rsTechnician.BOF Then
        SetEmpNO = LTrim(RTrim(Null2String(rsTechnician!EMPNO)))
    End If
End Function

Function SetCodeNO(XXX As String) As String
    Dim rsTechnician                                   As ADODB.Recordset
    Set rsTechnician = New ADODB.Recordset
    Set rsTechnician = gconDMIS.Execute("Select * from CSMS_Contractor where COMPANYname = '" & XXX & "'")
    If Not rsTechnician.EOF And Not rsTechnician.BOF Then
        SetCodeNO = LTrim(RTrim(Null2String(rsTechnician!Code)))
    End If
End Function

Sub PrintProductivityInExcel()
    Dim RSTECH                                         As New ADODB.Recordset
    Dim FLAT_TIME                                      As Double
    Dim PROD_TIME                                      As Double
    Dim cnt                                            As Integer

    Set RSTECH = gconDMIS.Execute("SELECT * FROM CSMS_VW_TECHNICIAN ORDER BY TECH_NAME")
    If Not (RSTECH.BOF And RSTECH.EOF) Then
        'lblNOTES.Visible = True
        cnt = 9
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "TECHNICIAN PRODUCTIVITY.XLS")
        Set xlSheet = xlBook.Worksheets(1)


        xlSheet.Cells(2, "B") = COMPANY_NAME
        xlSheet.Cells(3, "B") = COMPANY_ADDRESS
        xlSheet.Cells(5, "B") = "From " & TXTFrom.Value & " to " & txtTO.Value
        xlSheet.Cells(32, "L") = GENERAL_MANAGER

        Do While Not RSTECH.EOF
            xlSheet.Cells(cnt, "A") = Null2String(RSTECH!EMPNO)
            xlSheet.Cells(cnt, "B") = Null2String(RSTECH!TECH_NAME)

            Set rsREPOR = gconDMIS.Execute("SELECT CSMS_Repor.REP_OR, CSMS_Repor.TRANSTYPE, CSMS_Ro_Det.DET_HRS, CSMS_Ro_Det.HRSWRK, " & _
                                         " CSMS_Repor.DTE_COMP FROM CSMS_Repor INNER JOIN " & _
                                         " dbo.CSMS_Ro_Det ON dbo.CSMS_Repor.REP_OR = dbo.CSMS_Ro_Det.REP_OR AND " & _
                                         " dbo.CSMS_Repor.TRANSTYPE = dbo.CSMS_Ro_Det.TRANSTYPE " & _
                                         " Where dbo.CSMS_Repor.DTE_COMP BETWEEN '" & TXTFrom.Value & "' and '" & txtTO.Value & "' AND LTRIM(RTRIM(CSMS_RO_DET.TECHCODE)) = '" & LTrim(RTrim(RSTECH!Technician)) & "'")
            If Not (rsREPOR.BOF And rsREPOR.EOF) Then
                Do While Not rsREPOR.EOF
                    FLAT_TIME = FLAT_TIME + NumericVal(rsREPOR!DET_HRS)
                    PROD_TIME = PROD_TIME + NumericVal(rsREPOR!HRSWRK)

                    rsREPOR.MoveNext
                Loop
            Else
            End If

            Dim ATTEND_HR                              As Double
            Dim AVAILA_HR                              As Double

            ATTEND_HR = 0
            AVAILA_HR = 0
            xlSheet.Cells(cnt, "E") = FLAT_TIME
            xlSheet.Cells(cnt, "I") = PROD_TIME

            ATTEND_HR = ComputeAttendedHours(RSTECH!EMPNO)(0)
            AVAILA_HR = ComputeAttendedHours(RSTECH!EMPNO)(1)

            xlSheet.Cells(cnt, "G") = ATTEND_HR
            xlSheet.Cells(cnt, "K") = ATTEND_HR - PROD_TIME

            If PROD_TIME <= 0 Then
                xlSheet.Cells(cnt, "M") = 0
            Else
                xlSheet.Cells(cnt, "M") = (FLAT_TIME / PROD_TIME) * 100
            End If
            If ATTEND_HR <= 0 Then
                xlSheet.Cells(cnt, "O") = 0
            Else
                xlSheet.Cells(cnt, "O") = (FLAT_TIME / ATTEND_HR) * 100
            End If

            If PROD_TIME <= 0 Then
                xlSheet.Cells(cnt, "Q") = 0
            Else
                xlSheet.Cells(cnt, "Q") = (ATTEND_HR / PROD_TIME) * 100
            End If

            FLAT_TIME = 0
            PROD_TIME = 0

            cnt = cnt + 1
            If cnt > 28 Then
                xlApp.Visible = True
                Set xlApp = Nothing

                xlSheet.Cells(5, "B") = "From " & TXTFrom.Value & " to " & txtTO.Value
                xlSheet.Cells(2, "B") = COMPANY_NAME
                xlSheet.Cells(3, "B") = COMPANY_ADDRESS
                xlSheet.Cells(32, "L") = GENERAL_MANAGER
                cnt = 15
            End If

            RSTECH.MoveNext
        Loop

        xlApp.Visible = True
        Set xlApp = Nothing

        'LogAudit "V", "MONTHLY TIME CONTROL ANALYSIS - REPORTS ", txtFROM.Value & " - " & txtTO.Value
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "TECHNICIAN REPORT", "", "", "", "TECHNICIAN PRODUCTIVITY : " & cboTechnicianPerformance_Report & " : " & TXTFrom & " - " & txtTO, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
    End If

    Set RSTECH = Nothing
    Set rsREPOR = Nothing
End Sub

Sub FillCombo()
    Dim rsTMP                                          As New ADODB.Recordset
    Dim NEYM                                           As String

    cboTechnicianPerformance_Report.Clear
    cboTechnicianPerformance_Report.AddItem "ALL"
    Set rsTMP = gconDMIS.Execute("SELECT * FROM CSMS_vw_Technician order by tech_name")
    If Not (rsTMP.BOF And rsTMP.EOF) Then
        Do While Not rsTMP.EOF
            'NEYM = Null2String(RSTMP!lastname) & ", " & Null2String(RSTMP!Firstname) & " " & Left(Null2String(RSTMP!MIDDLENAME), 1) & "."

            cboTechnicianPerformance_Report.AddItem Null2String(rsTMP!TECH_NAME)
            rsTMP.MoveNext
        Loop
    End If
    cboTechnicianPerformance_Report.Text = "ALL"
    Set rsTMP = Nothing
End Sub

Sub FillComboContractor()
    Dim rsTMP                                          As New ADODB.Recordset
    Dim NEYM                                           As String

    cboTechnicianPerformance_Report.Clear
    cboTechnicianPerformance_Report.AddItem "ALL"
    Set rsTMP = gconDMIS.Execute("SELECT * FROM CSMS_CONTRACTOR order by COMPANYNAME")
    If Not (rsTMP.BOF And rsTMP.EOF) Then
        Do While Not rsTMP.EOF
            'NEYM = Null2String(RSTMP!lastname) & ", " & Null2String(RSTMP!Firstname) & " " & Left(Null2String(RSTMP!MIDDLENAME), 1) & "."

            cboTechnicianPerformance_Report.AddItem Null2String(rsTMP!CompanyName)
            rsTMP.MoveNext
        Loop
    End If

    cboTechnicianPerformance_Report.Text = "ALL"
    Set rsTMP = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "TECHNICIAN LABOR EFFICIENCY") = False Then Exit Sub
    'On Error GoTo Errorcode

    If TXTFrom.Value > txtTO.Value Then
        MsgBox "Invalid Range Format", vbInformation, "CSMS"
        TXTFrom.SetFocus
        Exit Sub
    End If

    If txtTO.Value < TXTFrom.Value Then
        MsgBox "Invalid Range Format", vbInformation, "CSMS"
        txtTO.SetFocus
        Exit Sub
    End If

    'If MsgBox("Print in Excel", vbQuestion + vbYesNo, "CSMS") = vbYes Then
    Call PrintProductivityInExcel
    Exit Sub
    'End If

    Dim VRANGE                                         As String
    Dim rsKUTO                                         As New ADODB.Recordset
    Screen.MousePointer = 11
    Set rsKUTO = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE DTE_COMP BETWEEN '" & TXTFrom.Value & "' AND '" & txtTO.Value & "'")
    VRANGE = "From " & TXTFrom.Value & " to " & txtTO.Value
    If Not (rsKUTO.BOF And rsKUTO.EOF) Then
        If Option1.Value = True Then
            If cboTechnicianPerformance_Report.Text = "ALL" Then
                rptTechnician.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptTechnician.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptTechnician.Formulas(2) = "PrintedBy = '" & LOGNAME & "'"
                rptTechnician.Formulas(3) = "VRANGE = '" & VRANGE & "'"
                rptTechnician.WindowTitle = "TECHNICIAN PRODUCTIVITY REPORT"

                PrintSQLReport rptTechnician, CSMS_REPORT_PATH & "Technician_Productivity.rpt", "{CSMS_REPOR.DTE_COMP} >= date(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {CSMS_REPOR.DTE_COMP} <= date(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1

                'LogAudit "V", "TECHNICIAN PRODUCTIVITY REPORT", txtFROM & "-" & txtTO
            Else
                rptTechnician.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptTechnician.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptTechnician.Formulas(2) = "PrintedBy = '" & LOGNAME & "'"
                rptTechnician.Formulas(3) = "VRANGE = '" & VRANGE & "'"
                rptTechnician.WindowTitle = "TECHNICIAN PRODUCTIVITY REPORT"

                PrintSQLReport rptTechnician, CSMS_REPORT_PATH & "Technician_Productivity.rpt", "{CSMS_REPOR.DTE_COMP} >= date(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {CSMS_REPOR.DTE_COMP} <= date(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ") and {CSMS_VW_TECHNICIAN.EMPNO} = '" & SetEmpNO(cboTechnicianPerformance_Report.Text) & "'", CSMS_REPORT_CONNECTION, 1

                'LogAudit "V", "TECHNICIAN PRODUCTIVITY REPORT", txtFROM & "-" & txtTO
            End If
            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "TECHNICIAN REPORT", "", "", "", "TECHNICIAN PRODUCTIVITY : " & cboTechnicianPerformance_Report & " " & TXTFrom & " " & txtTO, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        Else
            If cboTechnicianPerformance_Report.Text = "ALL" Then
                rptTechnician.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptTechnician.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptTechnician.Formulas(2) = "PrintedBy = '" & LOGNAME & "'"
                rptTechnician.Formulas(3) = "VRANGE = '" & VRANGE & "'"
                rptTechnician.WindowTitle = "TECHNICIAN PRODUCTIVITY REPORT"

                PrintSQLReport rptTechnician, CSMS_REPORT_PATH & "Contractor_Productivity.rpt", "{CSMS_REPOR.DTE_COMP} >= date(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {CSMS_REPOR.DTE_COMP} <= date(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ")", CSMS_REPORT_CONNECTION, 1

                'LogAudit "V", "TECHNICIAN PRODUCTIVITY REPORT", txtFROM & "-" & txtTO
            Else
                rptTechnician.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptTechnician.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                rptTechnician.Formulas(2) = "PrintedBy = '" & LOGNAME & "'"
                rptTechnician.Formulas(3) = "VRANGE = '" & VRANGE & "'"
                rptTechnician.WindowTitle = "TECHNICIAN PRODUCTIVITY REPORT"

                PrintSQLReport rptTechnician, CSMS_REPORT_PATH & "Contractor_Productivity.rpt", "{CSMS_REPOR.DTE_COMP} >= date(" & Year(TXTFrom.Value) & "," & Month(TXTFrom.Value) & "," & Day(TXTFrom.Value) & ") AND {CSMS_REPOR.DTE_COMP} <= date(" & Year(txtTO.Value) & "," & Month(txtTO.Value) & "," & Day(txtTO.Value) & ") and {csms_contractor.code} = '" & SetCodeNO(cboTechnicianPerformance_Report.Text) & "'", CSMS_REPORT_CONNECTION, 1

                'LogAudit "V", "TECHNICIAN PRODUCTIVITY REPORT", txtFROM & "-" & txtTO
            End If
            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("V", "TECHNICIAN REPORT", "", "", "", "CONTRACTOR PRODUCTIVITY : " & cboTechnicianPerformance_Report & " " & TXTFrom & " - " & txtTO, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        End If
    Else
        ShowNoRecord
    End If

    Set rsKUTO = Nothing
    Screen.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    TXTFrom.Value = firstDay(LOGDATE)
    txtTO.Value = LOGDATE
    Option1_Click
    Screen.MousePointer = 0
End Sub

Private Sub txtFrom_LostFocus()
    TXTFrom.Value = Format(TXTFrom.Value, "SHORT DATE")
End Sub

Private Sub txtTo_LostFocus()
    txtTO.Value = Format(txtTO.Value, "SHORT DATE")
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


VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmHRMSYTDProcessing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Year-To-Date Processing"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4560
   ForeColor       =   &H00D8E9EC&
   Icon            =   "YTDProcessing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   4560
   Begin VB.CheckBox chkNOBONUS 
      Caption         =   "No Bonus!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   1935
      Width           =   1965
   End
   Begin VB.CheckBox chkUpdateMidYear 
      Caption         =   "Update Mid Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   5
      Top             =   1155
      Width           =   2295
   End
   Begin VB.CheckBox chkUpdateBonus 
      Caption         =   "Update Bonus"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   7
      Top             =   1755
      Width           =   2295
   End
   Begin VB.CheckBox chkUpdate13thMonth 
      Caption         =   "Update 13th Month"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   6
      Top             =   1455
      Width           =   2295
   End
   Begin VB.CheckBox chkTerminated 
      Caption         =   "Process Year-to-Date for Terminated Employees"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   585
      Width           =   4455
   End
   Begin VB.CheckBox chkPrintDet 
      Caption         =   "Print Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1380
      TabIndex        =   4
      Top             =   855
      Width           =   1875
   End
   Begin VB.ComboBox cboDay 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   1125
   End
   Begin VB.ComboBox cboYear 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      Left            =   3420
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   135
      Width           =   1065
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   360
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   135
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2190
      MouseIcon       =   "YTDProcessing.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "YTDProcessing.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cancel"
      Top             =   2430
      Width           =   885
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   1320
      MouseIcon       =   "YTDProcessing.frx":079A
      MousePointer    =   99  'Custom
      Picture         =   "YTDProcessing.frx":08EC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Process Year-to-Date Processing"
      Top             =   2430
      Width           =   885
   End
   Begin Crystal.CrystalReport rptPrintYTD 
      Left            =   3600
      Top             =   1905
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
   Begin wizProgBar.Prg progYTDProcessing 
      Height          =   315
      Left            =   60
      TabIndex        =   11
      Top             =   3360
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   556
      Picture         =   "YTDProcessing.frx":0C5A
      BackColor       =   14215660
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "YTDProcessing.frx":0C76
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
   Begin VB.Label labEmpNo 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E37331&
      Height          =   225
      Left            =   90
      TabIndex        =   10
      Top             =   3090
      Width           =   2925
   End
End
Attribute VB_Name = "frmHRMSYTDProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSEMPINFO, rsPAYROLL, rsSalaryGrade                               As ADODB.Recordset
Attribute rsPAYROLL.VB_VarUserMemId = 1073938432
Attribute rsSalaryGrade.VB_VarUserMemId = 1073938432
Dim rsYTDDETAILS, rsCommission, rsvw_Adjustment                       As ADODB.Recordset
Attribute rsYTDDETAILS.VB_VarUserMemId = 1073938435
Attribute rsCommission.VB_VarUserMemId = 1073938435
Attribute rsvw_Adjustment.VB_VarUserMemId = 1073938435
Dim ToDate                                                            As String
Attribute ToDate.VB_VarUserMemId = 1073938438

Function getAssumedPH(GENDATE As Date, Salari As Double) As Double
    getAssumedPH = PhilHealthShare(Salari) * (12 - Month(GENDATE))
End Function

Function getAssumedSSS(GENDATE As Date, Salari As Double) As Double
    getAssumedSSS = EmployeeSSSshare(Salari) * (12 - Month(GENDATE))
End Function

Function getAssumedPagIbig(GENDATE As Date, Salari As Double) As Double
    getAssumedPagIbig = PagIbigShare(Salari) * (12 - Month(GENDATE))
End Function

Function SetSalary(SalCode As String) As Double
    Set rsSalaryGrade = New ADODB.Recordset
    rsSalaryGrade.Open "select code,salary from HRMS_SalaryGrade where code = '" & SalCode & "'", gconDMIS
    If Not rsSalaryGrade.EOF And Not rsSalaryGrade.BOF Then
        SetSalary = N2Str2Zero(rsSalaryGrade!Salary)
    End If
End Function

Function SetDailyRate(SalCode As String) As Double
    Set rsSalaryGrade = New ADODB.Recordset
    rsSalaryGrade.Open "select code,dailyrate from HRMS_SalaryGrade where code = '" & SalCode & "'", gconDMIS
    If Not rsSalaryGrade.EOF And Not rsSalaryGrade.BOF Then
        SetDailyRate = N2Str2Zero(rsSalaryGrade!DailyRate)
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    'If Function_Access(LOGID, "Acess_Process", "PROCESS YEAR-TO-DATE PROCESS") = False Then Exit Sub
    On Error GoTo Errorcode
    Dim MM, DD, YY                                                    As String
    MM = What_month(cboMonth)
    YY = cboYear.Text
    DD = cboDay.Text

    ToDate = DateSerial(YY, MM, DD)
    GENTO = Format(ToDate, "Short Date")
    Beep
    Dim i                                                             As Integer

    Dim VEMPNO, VTaxCode                                              As String
    Dim VYTDBasicPay, VYTDUTLate, VYTDAbsent                          As Double
    Dim VCommission, VCommissionTax, VDecCommissionTax                As Double
    Dim VOvertime, VTaxableAdj, VNonTaxableAdj                        As Double
    Dim Vsss, Vphic, Vpagibig, VCOLA, VYTDIncome                      As Double
    Dim VRemSal, VRemCOLA, VRemOT                                     As Double
    Dim VAccSalary, VAccCOLA, VAccBULANAN, VAccNONTAXABLE             As Double
    Dim VRemWTax, VCURRemSal, VCURRemOT                               As Double
    Dim VCURRemWTax, VMidYear, V13thMonth                             As Double
    Dim VPersonalEx, VYTDTaxable, VYTDNonTaxable                      As Double
    Dim VNetTaxable, VNetTax, VDecNetTax                              As Double
    Dim VTaxDue                                                       As Double

    Dim VTOTYTDBasicPay, VTOTCommission, VTOTCommissionTax            As Double
    Dim VDECTOTCommissionTax, VTOTOvertime, VTOTTaxableAdj            As Double
    Dim VTOTNonTaxableAdj, VTOTsss, VTOTphic                          As Double
    Dim VTOTpagibig, VTOTcola, VTOTYTDIncome, VTOTPersonalEx          As Double
    Dim VTOTYTDTaxable, VTOTYTDNonTaxable, VTOTNetTaxable             As Double
    Dim VTOTNetTax, VDECTOTNetTax, VTOTTaxDue, VBONUS                 As Double

    Dim VPAYempno, VPAYtaxcode                                        As String
    Dim VPAYrate, VPAYdailyrate, VPAYovertime                         As Double
    Dim VPAYAdjustment                                                As Double
    Dim VPAYcommission, VPAYcommissionTax, VDECPAYcommissionTax       As Double
    Dim VPAYtaxableadj, VPAYnontaxableadj, VPAYgross                  As Double
    Dim VPAYundertime, VPAYsss, VPAYphilhealth                        As Double
    Dim VPAYpagibig, VPAYcola, VPAYtin, VDECPAYtin                    As Double
    Dim VPAYabsent                                                    As Double

    Dim VTOTPAYempno, VTOTPAYtaxcode                                  As String
    Dim VTOTPAYrate, VTOTPAYdailyrate, VTOTPAYovertime                As Double
    Dim VTOTPAYAdjustment                                             As Double
    Dim VTOTPAYcommission, VTOTPAYcommissionTax, VDECTOTPAYcommissionTax As Double
    Dim VTOTPAYtaxableadj, VTOTPAYnontaxableadj, VTOTPAYgross         As Double
    Dim VTOTPAYundertime, VTOTPAYsss, VTOTPAYphilhealth               As Double
    Dim VTOTPAYpagibig, VTOTPAYcola, VTOTPAYtin, VDECTOTPAYtin        As Double
    Dim VTOTPAYabsent                                                 As Double

    Dim VNOVDECBASICPAY                                               As Double
    Dim VARYEER                                                       As String
    Dim NoMonths, manths, manths2                                     As Integer
    Dim VARYTDINCOME, VARPERSONALEX, BULANAN, COLA_RATE               As Double

    Dim CutOffDate                                                    As String
    CutOffDate = ""

    Call LogAudit("G", "GENERATE YEAR TO DATE PROCESSING", cboMonth & " " & cboDay & "," & cboYear)

    If chkTerminated.Value = 1 Then
        Set RSEMPINFO = New ADODB.Recordset
        RSEMPINFO.Open "select * from HRMS_EmpInfo where activeinactive = 'I' order by emplevel,empno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        Set RSEMPINFO = New ADODB.Recordset
        '    rsEMPINFO.Open "select * from HRMS_EmpInfo where activeinactive = 'A' order by emplevel,empno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        RSEMPINFO.Open "select * from HRMS_EmpInfo where activeinactive = 'A' and emplevel = 'E' order by emplevel,empno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If Not RSEMPINFO.EOF And Not RSEMPINFO.BOF Then
        RSEMPINFO.MoveFirst
        Do While Not RSEMPINFO.EOF
            If Null2String(RSEMPINFO!EMPNO) <> "" Then
                CutOffDate = ToDate
                VARYEER = YEAR(ToDate)
                VEMPNO = "": VTaxCode = "": NoMonths = 0
                VYTDBasicPay = 0: VCommission = 0: VOvertime = 0
                VTaxableAdj = 0: VNonTaxableAdj = 0: VCommissionTax = 0: VDecCommissionTax = 0
                VYTDIncome = 0: VYTDTaxable = 0: VYTDNonTaxable = 0
                VNetTaxable = 0: VNetTax = 0: VDecNetTax = 0: VAccSalary = 0
                VRemSal = 0: V13thMonth = 0: VRemOT = 0: VRemWTax = 0
                Vsss = 0: Vphic = 0: Vpagibig = 0: VCOLA = 0

                VPAYrate = 0: VPAYdailyrate = 0: VPAYcommissionTax = 0: VDECPAYcommissionTax = 0
                VPAYovertime = 0: VPAYcommission = 0: VPAYtaxableadj = 0: VDECPAYtin = 0: VPAYtin = 0
                VPAYnontaxableadj = 0: VPAYundertime = 0: VPAYsss = 0
                VPAYphilhealth = 0: VPAYpagibig = 0: VPAYcola = 0: VPAYabsent = 0

                VTOTPAYrate = 0: VTOTPAYdailyrate = 0: VTOTPAYtin = 0: VDECTOTPAYtin = 0
                VTOTPAYovertime = 0: VTOTPAYtaxableadj = 0: VMidYear = 0
                VTOTPAYnontaxableadj = 0: VTOTPAYundertime = 0: VTOTPAYsss = 0
                VTOTPAYphilhealth = 0: VTOTPAYpagibig = 0: VTOTPAYcola = 0: VTOTPAYabsent = 0
                VTOTPAYcommissionTax = 0: VDECTOTPAYcommissionTax = 0: VTOTPAYcommission = 0: VTOTPAYAdjustment = 0

                VPAYempno = Null2String(RSEMPINFO!EMPNO)
                COLA_RATE = N2Str2Zero(RSEMPINFO!COLA_RATE)
                VNOVDECBASICPAY = 0

                'PAYROLL *********************************************************************************
                Set rsPAYROLL = New ADODB.Recordset
                If Null2String(RSEMPINFO!EMPLEVEL) = "C" Then
                    rsPAYROLL.Open "select * from HRMS_Payroll where (EMPLEVEL = " & N2Str2Null(RSEMPINFO!EMPLEVEL) & ") AND year(paydateto) = " & YEAR(ToDate) & " AND empno =" & N2Str2Null(RSEMPINFO!EMPNO) & " order by paydateto desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                Else
                    rsPAYROLL.Open "select * from HRMS_Payroll where (EMPLEVEL = " & N2Str2Null(RSEMPINFO!EMPLEVEL) & ") AND year(paydateto) = " & YEAR(ToDate) & " AND (paydateto <= '" & Format(ToDate, "Short Date") & "') AND empno =" & N2Str2Null(RSEMPINFO!EMPNO) & " order by paydateto desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                End If
                If Not rsPAYROLL.EOF And Not rsPAYROLL.BOF Then
                    rsPAYROLL.MoveFirst
                    Do While Not rsPAYROLL.EOF
                        labEmpNo.Caption = 0
                        VPAYrate = 0: VPAYdailyrate = 0: VPAYtin = 0: VDECPAYtin = 0
                        VPAYovertime = 0: VPAYtaxableadj = 0
                        VPAYnontaxableadj = 0: VPAYundertime = 0: VPAYsss = 0
                        VPAYphilhealth = 0: VPAYpagibig = 0: VPAYcola = 0: VPAYabsent = 0
                        VPAYcommissionTax = 0: VDECPAYcommissionTax = 0: VPAYcommission = 0

                        manths = Month(rsPAYROLL!paydatefrom)
                        If manths2 <> manths Then
                            manths2 = manths
                            NoMonths = NoMonths + 1
                        End If
                        VPAYtaxcode = Null2String(RSEMPINFO!ExStatus)
                        VPAYrate = N2Str2Zero(rsPAYROLL!Rate)
                        VPAYcola = N2Str2Zero(rsPAYROLL!cola)
                        VPAYdailyrate = N2Str2Zero(rsPAYROLL!DailyRate)
                        VPAYovertime = NumericVal(N2Str2Zero(rsPAYROLL!OVERTIME)) + NumericVal(N2Str2Zero(rsPAYROLL!HOLIDAY))
                        If Month(rsPAYROLL!paydateto) = 12 Then
                            VDECPAYtin = N2Str2Zero(rsPAYROLL!tax)
                            VDECTOTPAYtin = VDECTOTPAYtin + VDECPAYtin
                        Else
                            VPAYtin = N2Str2Zero(rsPAYROLL!tax)
                            VTOTPAYtin = VTOTPAYtin + VPAYtin
                        End If
                        If Month(rsPAYROLL!paydateto) > 10 Then
                            VNOVDECBASICPAY = VNOVDECBASICPAY + VPAYrate
                        End If
                        VPAYtaxableadj = N2Str2Zero(rsPAYROLL!taxableadj)
                        VPAYnontaxableadj = N2Str2Zero(rsPAYROLL!nontaxableadj)
                        VPAYsss = N2Str2Zero(rsPAYROLL!sssE)
                        VPAYphilhealth = N2Str2Zero(rsPAYROLL!philhealthE)
                        VPAYpagibig = N2Str2Zero(rsPAYROLL!PAGIBIG)
                        VPAYundertime = N2Str2Zero(rsPAYROLL!UNDERTIME)
                        VPAYabsent = N2Str2Zero(rsPAYROLL!absent)

                        VTOTPAYrate = VTOTPAYrate + VPAYrate
                        VTOTPAYcola = VTOTPAYcola + VPAYcola
                        VTOTPAYdailyrate = VTOTPAYdailyrate + VPAYdailyrate
                        VTOTPAYovertime = VTOTPAYovertime + VPAYovertime

                        VTOTPAYtaxableadj = VTOTPAYtaxableadj + VPAYtaxableadj
                        VTOTPAYnontaxableadj = VTOTPAYnontaxableadj + VPAYnontaxableadj
                        VTOTPAYundertime = VTOTPAYundertime + VPAYundertime
                        VTOTPAYsss = VTOTPAYsss + VPAYsss
                        VTOTPAYphilhealth = VTOTPAYphilhealth + VPAYphilhealth
                        VTOTPAYpagibig = VTOTPAYpagibig + VPAYpagibig
                        VTOTPAYabsent = VTOTPAYabsent + VPAYabsent
                        rsPAYROLL.MoveNext
                    Loop
                End If
                'PAYROLL *********************************************************************************

                'ADJUSTMENT *********************************************************************************
                Set rsvw_Adjustment = New ADODB.Recordset
                If Null2String(RSEMPINFO!EMPLEVEL) = "C" Then
                    Set rsvw_Adjustment = gconDMIS.Execute("Select * from HRMS_vw_Adjustment where IncludedIN13thMonth = 'Y' and (EMPLEVEL = " & N2Str2Null(RSEMPINFO!EMPLEVEL) & ") AND year(deyt) = " & YEAR(ToDate) & " AND empno ='" & RSEMPINFO!EMPNO & "' order by deyt asc")
                Else
                    Set rsvw_Adjustment = gconDMIS.Execute("Select * from HRMS_vw_Adjustment where IncludedIN13thMonth = 'Y' and (EMPLEVEL = " & N2Str2Null(RSEMPINFO!EMPLEVEL) & ") AND year(deyt) = " & YEAR(ToDate) & " AND (deyt <= '" & Format(ToDate, "Short Date") & "') AND empno ='" & RSEMPINFO!EMPNO & "' order by deyt asc")
                End If
                If Not rsvw_Adjustment.EOF And Not rsvw_Adjustment.BOF Then
                    rsvw_Adjustment.MoveFirst: VPAYAdjustment = 0
                    Do While Not rsvw_Adjustment.EOF
                        VPAYAdjustment = N2Str2Zero(rsvw_Adjustment!AMOUNT)
                        If Month(rsvw_Adjustment!DEYT) = 12 Then
                            VTOTPAYAdjustment = VTOTPAYAdjustment + VPAYAdjustment
                        Else
                            VTOTPAYAdjustment = VTOTPAYAdjustment + VPAYAdjustment
                        End If
                        rsvw_Adjustment.MoveNext
                    Loop
                End If
                'ADJUSTMENT *********************************************************************************

                'COMMISSION *********************************************************************************
                Set rsCommission = New ADODB.Recordset
                If Null2String(RSEMPINFO!EMPLEVEL) = "C" Then
                    rsCommission.Open "select * from HRMS_Commission where (EMPLEVEL = " & N2Str2Null(RSEMPINFO!EMPLEVEL) & ") AND year(deyt) = " & YEAR(ToDate) & " AND empno ='" & RSEMPINFO!EMPNO & "' order by deyt asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                Else
                    rsCommission.Open "select * from HRMS_Commission where (EMPLEVEL = " & N2Str2Null(RSEMPINFO!EMPLEVEL) & ") AND year(deyt) = " & YEAR(ToDate) & " AND (deyt <= '" & Format(ToDate, "Short Date") & "') AND empno ='" & RSEMPINFO!EMPNO & "' order by deyt asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                End If
                If Not rsCommission.EOF And Not rsCommission.BOF Then
                    rsCommission.MoveFirst: VPAYcommissionTax = 0: VDECPAYcommissionTax = 0: VPAYcommission = 0
                    Do While Not rsCommission.EOF
                        VPAYcommission = N2Str2Zero(rsCommission!AMOUNT)
                        If Month(rsCommission!DEYT) = 12 Then
                            VTOTPAYcommission = VTOTPAYcommission + VPAYcommission
                            VDECPAYcommissionTax = N2Str2Zero(rsCommission!tax)
                            VDECTOTPAYcommissionTax = VDECTOTPAYcommissionTax + VDECPAYcommissionTax
                        Else
                            VTOTPAYcommission = VTOTPAYcommission + VPAYcommission
                            VPAYcommissionTax = N2Str2Zero(rsCommission!tax)
                            VTOTPAYcommissionTax = VTOTPAYcommissionTax + VPAYcommissionTax
                        End If
                        rsCommission.MoveNext
                    Loop
                End If
                'COMMISSION *********************************************************************************

                VEMPNO = VPAYempno
                VTaxCode = VPAYtaxcode
                VYTDBasicPay = (VTOTPAYrate) - (VTOTPAYundertime + VTOTPAYabsent)
                VCOLA = VTOTPAYcola
                VYTDUTLate = VTOTPAYundertime
                VYTDAbsent = VTOTPAYabsent
                VCommission = VTOTPAYcommission
                VCommissionTax = VTOTPAYcommissionTax
                VDecCommissionTax = VDECTOTPAYcommissionTax
                VOvertime = VTOTPAYovertime
                VTaxableAdj = VTOTPAYtaxableadj
                VNonTaxableAdj = VTOTPAYnontaxableadj
                VARYTDINCOME = (VTOTPAYrate + VCOLA + VTaxableAdj + VCommission + VTOTPAYovertime) - (VTOTPAYundertime + VTOTPAYabsent)
                VARPERSONALEX = Personal_EX(VPAYtaxcode)
                Vsss = VTOTPAYsss
                Vphic = VTOTPAYphilhealth
                Vpagibig = VTOTPAYpagibig
                If Null2String(RSEMPINFO!EMPSTATUS) = "M" Then
                    BULANAN = SetSalary(Null2String(RSEMPINFO!SalaryCode))
                Else
                    BULANAN = (SetDailyRate(Null2String(RSEMPINFO!SalaryCode)) * 314) / 12
                End If

                Set rsYTDDETAILS = New ADODB.Recordset
                rsYTDDETAILS.Open "select * from HRMS_ytddetails where (EMPLEVEL = " & N2Str2Null(RSEMPINFO!EMPLEVEL) & ") AND empno = '" & VEMPNO & "' and yeer = '" & VARYEER & "'", gconDMIS, adOpenKeyset, adLockOptimistic
                If Not rsYTDDETAILS.EOF And Not rsYTDDETAILS.BOF Then

                    VMidYear = N2Str2Zero(rsYTDDETAILS!midyear)
                    VCURRemSal = N2Str2Zero(rsYTDDETAILS!remsal)
                    VCURRemOT = N2Str2Zero(rsYTDDETAILS!remot)
                    VCURRemWTax = N2Str2Zero(rsYTDDETAILS!remwtax)
                    VBONUS = N2Str2Zero(rsYTDDETAILS!bonus)
                    VRemCOLA = N2Str2Zero(rsYTDDETAILS!RemCOLA)
                    VRemSal = VCURRemSal
                    VRemOT = VRemOT
                    VRemWTax = VRemWTax + (VCURRemWTax - VRemWTax)

                    '================================================================
                    If Null2String(RSEMPINFO!RESIGNED) = "" Then
                        If Null2String(RSEMPINFO!ACTIVEINACTIVE) = "A" Then
                            If Null2String(RSEMPINFO!EMPLEVEL) <> "C" Then
                                If Month(CDate(CutOffDate)) = 4 Then
                                    VAccSalary = BULANAN * 2
                                    VMidYear = (((VYTDBasicPay + VTOTPAYAdjustment) + VAccSalary)) / 12
                                ElseIf Month(CDate(CutOffDate)) = 5 Then
                                    VAccSalary = BULANAN
                                    VMidYear = (((VYTDBasicPay + VTOTPAYAdjustment) + VAccSalary)) / 12
                                ElseIf Month(CDate(CutOffDate)) = 10 Then
                                    VAccSalary = BULANAN * 2
                                ElseIf Month(CDate(CutOffDate)) = 11 Then
                                    VAccSalary = BULANAN
                                Else
                                    VAccSalary = 0
                                End If
                                COLA_RATE = COLA_RATE * 26.17
                                VAccBULANAN = BULANAN
                                VAccCOLA = COLA_RATE * (12 - Month(CDate(CutOffDate)))
                                VAccBULANAN = VAccBULANAN * (12 - Month(CDate(CutOffDate)))
                            Else
                                VAccSalary = 0: VAccCOLA = 0: VAccBULANAN = 0
                                VMidYear = (((VYTDBasicPay + VTOTPAYAdjustment) + VAccSalary)) / 12
                            End If
                        Else
                            VAccSalary = 0: VAccCOLA = 0: VAccBULANAN = 0
                        End If

                        V13thMonth = Round((((VYTDBasicPay + VTOTPAYAdjustment) + VAccSalary)) / 12, 2)
                        If chkNOBONUS.Value = 1 Then
                            VBONUS = 0
                        Else
                            VBONUS = V13thMonth / 4
                        End If

                        If Month(CDate(CutOffDate)) > 6 Then
                            V13thMonth = V13thMonth - VMidYear
                        End If
                    Else
                        VAccSalary = 0: VAccCOLA = 0: VAccBULANAN = 0
                        V13thMonth = (((VYTDBasicPay + VTOTPAYAdjustment) + VRemSal)) / 12
                    End If
                    '================================================================

                    'UPDATE DEC. 29, 2005
                    If chkUpdateMidYear.Value = 0 Then VMidYear = N2Str2Zero(rsYTDDETAILS!midyear)
                    If chkUpdate13thMonth.Value = 0 Then V13thMonth = N2Str2Zero(rsYTDDETAILS!t13thmonth)
                    If chkUpdateBonus.Value = 0 Then VBONUS = N2Str2Zero(rsYTDDETAILS!bonus)
                    If V13thMonth + VMidYear + VBONUS > 30000 Then VARYTDINCOME = VARYTDINCOME + ((V13thMonth + VMidYear + VBONUS) - 30000)
                    'END UPDATE OF DEC. 29, 2005

                    VYTDTaxable = (VARYTDINCOME + VAccBULANAN + VAccCOLA + VRemCOLA + VRemOT + VRemSal)
                    VAccNONTAXABLE = getAssumedSSS(CDate(CutOffDate), NumericVal(BULANAN)) + getAssumedPH(CDate(CutOffDate), NumericVal(BULANAN)) + getAssumedPagIbig(CDate(CutOffDate), NumericVal(BULANAN))
                    VYTDNonTaxable = VTOTPAYsss + VTOTPAYphilhealth + VTOTPAYpagibig
                    VYTDBasicPay = (VYTDBasicPay + VOvertime) - VYTDNonTaxable
                    VNetTaxable = VYTDTaxable - (VYTDNonTaxable + VARPERSONALEX + (VAccNONTAXABLE))
                    VNetTax = VTOTPAYtin + VRemWTax
                    VDecNetTax = VDECTOTPAYtin

                    If VYTDTaxable >= 60000 Then
                        VTaxDue = Tax_Due(VNetTaxable)
                    Else
                        VTaxDue = 0
                    End If
                    'VRemSal = VAccSalary
                    gconDMIS.Execute "update HRMS_YTDdetails set " & _
                                     "taxcode = '" & VTaxCode & "', " & _
                                     "YTDGross = " & VTOTPAYrate & ", ytdbasicpay = " & VYTDBasicPay & ", " & _
                                     "ytdutlate = " & VYTDUTLate & ", " & _
                                     "ytdabsent = " & VYTDAbsent & ", " & _
                                     "Adjustment = " & VTOTPAYAdjustment & ", commission = " & VCommission & ", " & _
                                     "commissiontax = " & VCommissionTax & ", DecCommissiontax = " & VDecCommissionTax & ", " & _
                                     "overtime =" & VOvertime & ", " & _
                                     "taxableadj = " & VTaxableAdj & ", " & _
                                     "nontaxableadj = " & VNonTaxableAdj & ", " & _
                                     "ytdsss = " & Vsss & ", " & _
                                     "ytdphic = " & Vphic & ", ytdpagibig = " & Vpagibig & ", ytdcola = " & VCOLA & ", " & _
                                     "YTDincome = " & VARYTDINCOME + VRemSal + VRemOT & ", " & _
                                     "personalex = " & VARPERSONALEX & ", " & _
                                     "ytdtaxable = " & VYTDTaxable & ", " & _
                                     "nontaxable = " & VYTDNonTaxable & ", " & _
                                     "nettaxable = " & VNetTaxable & ", " & _
                                     "ytdtax = " & VNetTax & ", decytdtax = " & VDecNetTax & ", " & _
                                     "remcola = " & VRemCOLA & ", remsal = " & VRemSal & ", remot = " & VRemOT & ", remwtax = " & VRemWTax & ", " & _
                                     "DateHired = " & N2Date2Null(RSEMPINFO!DateHired) & ", " & _
                                     "ytdcutoffdate = " & N2Date2Null(CutOffDate) & ", " & _
                                     "ytdgenerate = " & N2Date2Null(GENTO) & ", " & _
                                     "midyear = " & VMidYear & ", Bonus = " & VBONUS & ", t13thmonth = " & V13thMonth & ", " & _
                                     "taxdue = " & VTaxDue & _
                                   " where (EMPLEVEL = " & N2Str2Null(RSEMPINFO!EMPLEVEL) & ") AND empno = '" & VEMPNO & "' AND YEER = '" & VARYEER & "'"
                    If chkUpdateMidYear.Value = 1 Then
                        gconDMIS.Execute "update HRMS_YTDdetails set " & _
                                         "midyear = " & VMidYear & _
                                       " where (EMPLEVEL = " & N2Str2Null(RSEMPINFO!EMPLEVEL) & ") AND empno = '" & VEMPNO & "' AND YEER = '" & VARYEER & "'"
                    End If

                    If chkUpdate13thMonth.Value = 1 Then
                        gconDMIS.Execute "update HRMS_YTDdetails set " & _
                                         "t13thmonth = " & V13thMonth & _
                                       " where (EMPLEVEL = " & N2Str2Null(RSEMPINFO!EMPLEVEL) & ") AND empno = '" & VEMPNO & "' AND YEER = '" & VARYEER & "'"
                    End If

                    If chkUpdateBonus.Value = 1 Then
                        gconDMIS.Execute "update HRMS_YTDdetails set " & _
                                         "Bonus = " & VBONUS & _
                                       " where (EMPLEVEL = " & N2Str2Null(RSEMPINFO!EMPLEVEL) & ") AND empno = '" & VEMPNO & "' AND YEER = '" & VARYEER & "'"
                    End If

                    gconDMIS.Execute "update HRMS_YTDdetails set " & _
                                     "taxcode = '" & VTaxCode & "', " & _
                                     "YTDGross = " & VTOTPAYrate & ", ytdbasicpay = " & VYTDBasicPay & ", " & _
                                     "ytdutlate = " & VYTDUTLate & ", " & _
                                     "ytdabsent = " & VYTDAbsent & ", " & _
                                     "Adjustment = " & VTOTPAYAdjustment & ", commission = " & VCommission & ", " & _
                                     "commissiontax = " & VCommissionTax & ", DecCommissiontax = " & VDecCommissionTax & ", " & _
                                     "overtime =" & VOvertime & ", " & _
                                     "taxableadj = " & VTaxableAdj & ", " & _
                                     "nontaxableadj = " & VNonTaxableAdj & ", " & _
                                     "ytdsss = " & Vsss & ", " & _
                                     "ytdphic = " & Vphic & ", ytdpagibig = " & Vpagibig & ", ytdcola = " & VCOLA & ", " & _
                                     "YTDincome = " & VARYTDINCOME + VRemSal + VRemOT & ", " & _
                                     "personalex = " & VARPERSONALEX & ", " & _
                                     "ytdtaxable = " & VYTDTaxable & ", " & _
                                     "nontaxable = " & VYTDNonTaxable & ", " & _
                                     "nettaxable = " & VNetTaxable & ", " & _
                                     "ytdtax = " & VNetTax & ", decytdtax = " & VDecNetTax & ", " & _
                                     "remcola = " & VRemCOLA & ", remsal = " & VRemSal & ", remot = " & VRemOT & ", remwtax = " & VRemWTax & ", " & _
                                     "DateHired = " & N2Date2Null(RSEMPINFO!DateHired) & ", " & _
                                     "ytdcutoffdate = " & N2Date2Null(CutOffDate) & ", " & _
                                     "ytdgenerate = " & N2Date2Null(GENTO) & ", " & _
                                     "t13thmonth = " & V13thMonth & ", " & _
                                     "taxdue = " & VTaxDue & _
                                   " where (EMPLEVEL = " & N2Str2Null(RSEMPINFO!EMPLEVEL) & ") AND empno = '" & VEMPNO & "' AND YEER = '" & VARYEER & "'"
                Else
                    If Null2String(RSEMPINFO!RESIGNED) = "" Then
                        If Null2String(RSEMPINFO!ACTIVEINACTIVE) = "A" Then
                            If Month(CDate(CutOffDate)) = 4 Then
                                VAccSalary = BULANAN * 2
                                VMidYear = (((VYTDBasicPay - VOvertime) + VAccSalary)) / 12
                            ElseIf Month(CDate(CutOffDate)) = 5 Then
                                VAccSalary = BULANAN
                                VMidYear = (((VYTDBasicPay - VOvertime) + VAccSalary)) / 12
                            ElseIf Month(CDate(CutOffDate)) = 10 Then
                                VAccSalary = BULANAN * 2
                            ElseIf Month(CDate(CutOffDate)) = 11 Then
                                VAccSalary = BULANAN
                            Else
                                VAccSalary = 0
                            End If
                        Else
                            VAccSalary = 0
                        End If

                        V13thMonth = Round((((VYTDBasicPay) + VAccSalary)) / 12, 2)
                        If Month(CDate(CutOffDate)) > 6 Then
                            V13thMonth = Round(V13thMonth - VMidYear, 2)
                        End If
                    Else
                        V13thMonth = Round((((VYTDBasicPay) + VRemSal)) / 12, 2)
                    End If

                    VYTDTaxable = Round((VARYTDINCOME + VRemOT + VRemSal) - VARPERSONALEX, 2)
                    VYTDNonTaxable = Round(VTOTPAYsss + VTOTPAYphilhealth + VTOTPAYpagibig, 2)
                    VYTDBasicPay = Round((VYTDBasicPay + VTOTPAYovertime) - VYTDNonTaxable, 2)
                    VNetTaxable = Round(VYTDTaxable - VYTDNonTaxable, 2)
                    VNetTax = Round(VTOTPAYtin + VRemWTax, 2)
                    VDecNetTax = Round(VDECTOTPAYtin, 2)
                    VTaxDue = Round(Tax_Due(VNetTaxable), 2)
                    gconDMIS.Execute "insert into HRMS_YTDdetails " & _
                                     "(EMPLEVEL,yeer,empno,taxcode,ytdbasicpay,ytdutlate,ytdabsent,commission,commissiontax,deccommissiontax,overtime" & _
                                     ",taxableadj,nontaxableadj,YTDincome,personalex" & _
                                     ",ytdtaxable,nontaxable,nettaxable,ytdtax,decytdtax,taxdue,remsal,midyear,t13thmonth,datehired,ytdcutoffdate,ytdgenerate)" & _
                                   " values (" & N2Str2Null(RSEMPINFO!EMPLEVEL) & ",'" & VARYEER & "', '" & VEMPNO & "', '" & VTaxCode & _
                                     "', " & VYTDBasicPay & ", " & VYTDUTLate & ", " & VYTDAbsent & ", " & VCommission & ", " & VCommissionTax & ", " & VDecCommissionTax & _
                                     ", " & VOvertime & ", " & VTaxableAdj & _
                                     ", " & VNonTaxableAdj & _
                                     ", " & VARYTDINCOME & ", " & VARPERSONALEX & _
                                     ", " & VYTDTaxable & ", " & VYTDNonTaxable & _
                                     ", " & VNetTaxable & ", " & VNetTax & ", " & VDecNetTax & ", " & VTaxDue & ", " & VRemSal & ", " & VMidYear & ", " & V13thMonth & ", " & N2Date2Null(RSEMPINFO!DateHired) & ", " & N2Date2Null(CutOffDate) & ", " & N2Date2Null(GENTO) & ")"
                End If
            End If
            i = i + 1
            progYTDProcessing.Value = (i / RSEMPINFO.RecordCount) * 100
            labEmpNo.Caption = Int(progYTDProcessing.Value) & "%"
            DoEvents
            RSEMPINFO.MoveNext
        Loop
    Else
        MsgBox "NO RECORD!"
    End If
    If LEDGERSHOW = True Then
        frmHRMSLedger.rsRefresh
        frmHRMSLedger.storeMemvars
        DoEvents
    End If
    If chkPrintDet.Value = 1 Then
        Screen.MousePointer = 11
        PrintSQLReport rptPrintYTD, HRMS_REPORT_PATH & "ytddetails.rpt", "", DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    End If
    Screen.MousePointer = 0
    Exit Sub

Errorcode:
    Screen.MousePointer = 0
    MsgBoxXP "Error: " & Err.NUMBER & vbCrLf & "Description: " & Err.Description, "Error", XP_OKOnly, msg_Critical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    fillcboDay cboDay
    fillcbomonth cboMonth
    FillcboYear cboYear
    If Day(LOGDATE) > 15 Then
        cboDay.Text = Day(lastDay(LOGDATE))
    Else
        cboDay.Text = 15
    End If
    cboYear.Text = YEAR(LOGDATE)
    cboMonth.Text = The_month(Month(LOGDATE))
    labEmpNo.Caption = ""
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub


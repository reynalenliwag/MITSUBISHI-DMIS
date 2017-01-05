VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRMSDTRSummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DTR Summary Report"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3675
   ForeColor       =   &H00D8E9EC&
   Icon            =   "DTRSummary.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2250
   ScaleWidth      =   3675
   Begin MSComctlLib.ProgressBar prg 
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   1890
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
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
      Height          =   795
      Left            =   2010
      MouseIcon       =   "DTRSummary.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "DTRSummary.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close Window"
      Top             =   1020
      Width           =   855
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
      Height          =   795
      Left            =   1170
      MouseIcon       =   "DTRSummary.frx":08A7
      MousePointer    =   99  'Custom
      Picture         =   "DTRSummary.frx":09F9
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print Report"
      Top             =   1020
      Width           =   855
   End
   Begin VB.CheckBox chkPreview 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   2010
      TabIndex        =   3
      Top             =   4020
      Width           =   1275
   End
   Begin VB.ComboBox cboQuensina 
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
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3465
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
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   540
      Width           =   2205
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
      Left            =   2370
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   540
      Width           =   1215
   End
   Begin Crystal.CrystalReport rptPrintPay 
      Left            =   1980
      Top             =   1080
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
End
Attribute VB_Name = "frmHRMSDTRSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim matt                                                              As String
Dim OT_REG                                                            As Double
Dim OT_ND_REG                                                         As Double
Dim SUN                                                               As Double
Dim OT_SUN                                                            As Double
Dim OT_ND_SUN                                                         As Double
Dim HOL_REG                                                           As Double
Dim HOL_REG_OT                                                        As Double
Dim HOL_REG_ND                                                        As Double
Dim HOL_S                                                             As Double
Dim HOL_S_OT                                                          As Double
Dim HOL_SND                                                           As Double
Dim HOL_SR                                                            As Double
Dim HOL_SR_OT                                                         As Double
Dim HOL_SRND                                                          As Double

Function GetOT(XEMPNO As String, OTCODE As String) As Double
    Dim rsot                                                          As ADODB.Recordset
    GetOT = 0
    Set rsot = gconDMIS.Execute("SELECT ISNULL(SUM(TOTALHR), 0) AS TOTALHOUR  FROM HRMS_OVERTIME" & _
                              " WHERE CUT_OFF = " & matt & _
                              " AND PAY_MONTH = '" & What_month(cboMOnth.Text) & "'" & _
                              " AND PAY_YEAR = '" & cboyear.Text & "'" & _
                              " AND EMPNO = '" & XEMPNO & "'" & _
                              " AND OCODE = '" & OTCODE & "'")

    GetOT = rsot!TOTALHOUR
    Set rsot = Nothing
End Function

Function GetTardiness(XEMPNO As String) As Integer
    Dim rsTardiness                                                   As ADODB.Recordset
    GetTardiness = 0
    Set rsTardiness = gconDMIS.Execute("SELECT ISNULL(SUM(NOMIN), 0) AS  MINUTESUM  FROM HRMS_DEDUCTIONS" & _
                                     " WHERE CUT_OFF = " & matt & _
                                     " AND PAY_MONTH = '" & What_month(cboMOnth.Text) & "'" & _
                                     " AND PAY_YEAR = '" & cboyear.Text & "'" & _
                                     " AND EMPNO = '" & XEMPNO & "'" & _
                                     " AND ISNULL(NOMIN, 0) > 0" & _
                                     " AND (PARTICULAR = 'LT' OR PARTICULAR = 'UT')")

    GetTardiness = rsTardiness!MINUTESUM
    Set rsTardiness = Nothing
End Function

Function GetTardinessTimes(XEMPNO As String) As Integer
    Dim rsTardiness                                                   As ADODB.Recordset
    GetTardinessTimes = 0
    Set rsTardiness = gconDMIS.Execute("SELECT COUNT(*) AS COUNTTIMES FROM HRMS_DEDUCTIONS" & _
                                     " WHERE CUT_OFF = " & matt & _
                                     " AND PAY_MONTH = '" & What_month(cboMOnth.Text) & "'" & _
                                     " AND PAY_YEAR = '" & cboyear.Text & "'" & _
                                     " AND EMPNO = '" & XEMPNO & "'" & _
                                     " AND ISNULL(NOMIN, 0) > 0" & _
                                     " AND (PARTICULAR = 'LT' OR PARTICULAR = 'UT')")

    GetTardinessTimes = rsTardiness!COUNTTIMES
    Set rsTardiness = Nothing
End Function

Function GetAbsencesTimes(XEMPNO As String) As Double
    Dim rsAbsences                                                    As ADODB.Recordset
    GetAbsencesTimes = 0
    Set rsAbsences = gconDMIS.Execute("SELECT ISNULL(SUM(NOMIN), 0) AS  MINUTESUM  FROM HRMS_DEDUCTIONS" & _
                                    " WHERE CUT_OFF = " & matt & _
                                    " AND PAY_MONTH = '" & What_month(cboMOnth.Text) & "'" & _
                                    " AND PAY_YEAR = '" & cboyear.Text & "'" & _
                                    " AND EMPNO = '" & XEMPNO & "'" & _
                                    " AND ISNULL(NOMIN, 0) > 0" & _
                                    " AND (PARTICULAR = 'WD' OR PARTICULAR = 'HD')")

    GetAbsencesTimes = rsAbsences!MINUTESUM / 480
    Set rsAbsences = Nothing
End Function

Function GetAbsencesDates(XEMPNO As String) As String
    Dim rsAbsencesDates                                               As ADODB.Recordset
    GetAbsencesDates = ""
    Set rsAbsencesDates = gconDMIS.Execute("SELECT * FROM HRMS_DEDUCTIONS" & _
                                         " WHERE CUT_OFF = " & matt & _
                                         " AND PAY_MONTH = '" & What_month(cboMOnth.Text) & "'" & _
                                         " AND PAY_YEAR = '" & cboyear.Text & "'" & _
                                         " AND EMPNO = '" & XEMPNO & "'" & _
                                         " AND ISNULL(NOMIN, 0) > 0" & _
                                         " AND (PARTICULAR = 'WD' OR PARTICULAR = 'HD') ORDER BY DEYT ASC")
    If Not rsAbsencesDates.EOF And Not rsAbsencesDates.BOF Then
        GetAbsencesDates = "LWOP "
        While Not rsAbsencesDates.EOF
            GetAbsencesDates = GetAbsencesDates + Null2String(Day(rsAbsencesDates!DEYT)) + ", "
            rsAbsencesDates.MoveNext
        Wend
    End If
    Set rsAbsencesDates = Nothing
End Function

Function GetOTDates(XEMPNO As String) As String
    Dim rsOTDates                                                     As ADODB.Recordset
    GetOTDates = ""
    Set rsOTDates = gconDMIS.Execute("SELECT * FROM HRMS_OVERTIME" & _
                                   " WHERE CUT_OFF = " & matt & _
                                   " AND PAY_MONTH = '" & What_month(cboMOnth.Text) & "'" & _
                                   " AND PAY_YEAR = '" & cboyear.Text & "'" & _
                                   " AND EMPNO = '" & XEMPNO & "' ORDER BY DEYT ASC")
    If Not rsOTDates.EOF And Not rsOTDates.BOF Then
        GetOTDates = "OT "
        While Not rsOTDates.EOF
            GetOTDates = GetOTDates + Null2String(Day(rsOTDates!DEYT)) + ", "
            rsOTDates.MoveNext
        Wend
    End If
    Set rsOTDates = Nothing
End Function

Function GetLeaveDates(XEMPNO As String, XTYPE As String) As String
    Dim rsLeaveDates                                                  As ADODB.Recordset
    GetLeaveDates = ""
    Set rsLeaveDates = gconDMIS.Execute("SELECT * FROM HRMS_REQUESTLEAVE_OT" & _
                                      " WHERE CUT_OFF = " & matt & _
                                      " AND PAY_MONTH = '" & What_month(cboMOnth.Text) & "'" & _
                                      " AND PAY_YEAR = '" & cboyear.Text & "'" & _
                                      " AND EMPNO = '" & XEMPNO & _
                                      "' AND REQCODE = '" & XTYPE & _
                                      "' AND STATUS = 'A'")
    If Not rsLeaveDates.EOF And Not rsLeaveDates.BOF Then
        GetLeaveDates = XTYPE
        While Not rsLeaveDates.EOF
            GetLeaveDates = GetLeaveDates + Null2String(Day(rsLeaveDates!DTE_FROM)) + ", "
            rsLeaveDates.MoveNext
        Wend
    End If
    Set rsLeaveDates = Nothing
End Function

Function HOUR_MIN(XHours) As String
    Dim strMin, strHour
    strMin = (XHours * 60) Mod 60
    strHour = ((XHours * 60) - ((XHours * 60) Mod 60)) \ 60
    If strMin = 0 And strHour = 0 Then
        Exit Function
    End If
    HOUR_MIN = strHour & "/" & Round(strMin, 2)
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Function GetSickLeave(XEMPNO As String) As Double
    Dim RSTMP As New ADODB.Recordset
    Dim SL_CNT As Double
    Dim date_from           As String
    Dim date_to             As String
    
    'If XEMPNO = "200702" Then Stop
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_RequestLeave_OT WHERE EMPNO = '" & XEMPNO & _
        "' AND CUT_OFF = " & matt & _
        " AND PAY_MONTH = " & What_month(cboMOnth) & _
        " AND PAY_YEAR = " & cboyear & _
        " AND REQCODE = 'SL'  AND STATUS = 'A'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        
        date_from = Trim(RSTMP!DTE_FROM)
        date_to = Trim(RSTMP!dte_to)
        
        
        Do While Not RSTMP.EOF
            If Hour(Null2String(RSTMP!OT_FROM)) = 8 Then
                If Hour(Null2String(RSTMP!OT_TO)) = 17 Then
                    
                    SL_CNT = (DateDiff("D", date_from, date_to) + 1)
                    'jbf
                    'SL_CNT = SL_CNT + 1
                Else
                    SL_CNT = SL_CNT + 0.5
                End If
            Else
                SL_CNT = SL_CNT + 0.5
            End If
            
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
    GetSickLeave = SL_CNT
End Function

Function GetMealTranspoAllowance(XEMPNO As String, XTYPE As String) As Double
    Dim RSTMP As New ADODB.Recordset
    Dim X_AMOUNT As Double
    Set RSTMP = gconDMIS.Execute("SELECT AMOUNT FROM HRMS_ADJUSTMENT WHERE PARTICULAR = '" & XTYPE & "' AND EMPNO = '" & XEMPNO & _
        "' AND CUT_OFF = " & matt & _
        " AND PAY_MONTH = " & What_month(cboMOnth) & _
        " AND PAY_YEAR = " & cboyear & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            X_AMOUNT = X_AMOUNT + NumericVal(RSTMP!AMOUNT)
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
    GetMealTranspoAllowance = X_AMOUNT
End Function

Function GetVacationLeave(XEMPNO As String) As Double
    Dim RSTMP               As New ADODB.Recordset
    Dim VL_CNT              As Double
    Dim date_from           As String
    Dim date_to             As String
    
    
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_RequestLeave_OT WHERE EMPNO = '" & XEMPNO & _
        "' AND CUT_OFF = " & matt & _
        " AND PAY_MONTH = " & What_month(cboMOnth) & _
        " AND PAY_YEAR = " & cboyear & _
        " AND REQCODE = 'VL'  AND STATUS = 'A'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        
        date_from = Trim(RSTMP!DTE_FROM)
        date_to = Trim(RSTMP!dte_to)
        
        
        Do While Not RSTMP.EOF
            
            If Hour(Null2String(RSTMP!OT_FROM)) = 8 Then
                If Hour(Null2String(RSTMP!OT_TO)) = 17 Then
                    
                    VL_CNT = (DateDiff("D", date_from, date_to) + 1)
                    'UPDATE: JBF
                    'VL_CNT = VL_CNT + 1
                Else
                    VL_CNT = VL_CNT + 0.5
                End If
            Else
                VL_CNT = VL_CNT + 0.5
            End If
            
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
    GetVacationLeave = VL_CNT
End Function

Function GetPaternityLeave(XEMPNO As String) As Double
    Dim RSTMP As New ADODB.Recordset
    Dim PL_CNT As Double
    Dim date_from           As String
    Dim date_to             As String
    
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_RequestLeave_OT WHERE EMPNO = '" & XEMPNO & _
        "' AND CUT_OFF = " & matt & _
        " AND PAY_MONTH = " & What_month(cboMOnth) & _
        " AND PAY_YEAR = " & cboyear & _
        " AND REQCODE = 'PL' AND STATUS = 'A'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        
        date_from = Trim(RSTMP!DTE_FROM)
        date_to = Trim(RSTMP!dte_to)
        
        Do While Not RSTMP.EOF
            If Hour(Null2String(RSTMP!OT_FROM)) = 8 Then
                If Hour(Null2String(RSTMP!OT_TO)) = 17 Then
                    
                    PL_CNT = (DateDiff("D", date_from, date_to) + 1)
                    'PL_CNT = PL_CNT + 1
                Else
                    PL_CNT = PL_CNT + 0.5
                End If
            Else
                PL_CNT = PL_CNT + 0.5
            End If
            
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
    GetPaternityLeave = PL_CNT
End Function

Function GetMaternityLeave(XEMPNO As String) As Double
    Dim RSTMP As New ADODB.Recordset
    Dim ML_CNT As Double
    Dim date_from           As String
    Dim date_to             As String
    
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_RequestLeave_OT WHERE EMPNO = '" & XEMPNO & _
        "' AND CUT_OFF = " & matt & _
        " AND PAY_MONTH = " & What_month(cboMOnth) & _
        " AND PAY_YEAR = " & cboyear & _
        " AND REQCODE = 'ML' AND STATUS = 'A'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        
        date_from = Trim(RSTMP!DTE_FROM)
        date_to = Trim(RSTMP!dte_to)
        
        Do While Not RSTMP.EOF
            If Hour(Null2String(RSTMP!OT_FROM)) = 8 Then
                If Hour(Null2String(RSTMP!OT_TO)) = 17 Then
                    
                    ML_CNT = (DateDiff("D", date_from, date_to) + 1)
                    'PL_CNT = PL_CNT + 1
                Else
                    ML_CNT = ML_CNT + 0.5
                End If
            Else
                ML_CNT = ML_CNT + 0.5
            End If
            
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
    GetMaternityLeave = ML_CNT
End Function

Function GetemergencyLeave(XEMPNO As String) As Double
    Dim RSTMP As New ADODB.Recordset
    Dim EL_CNT As Double
    Dim date_from           As String
    Dim date_to             As String
    
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_RequestLeave_OT WHERE EMPNO = '" & XEMPNO & _
        "' AND CUT_OFF = " & matt & _
        " AND PAY_MONTH = " & What_month(cboMOnth) & _
        " AND PAY_YEAR = " & cboyear & _
        " AND REQCODE = 'EL' AND STATUS = 'A'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        
        date_from = Trim(RSTMP!DTE_FROM)
        date_to = Trim(RSTMP!dte_to)
        
        Do While Not RSTMP.EOF
            If Hour(Null2String(RSTMP!OT_FROM)) = 8 Then
                If Hour(Null2String(RSTMP!OT_TO)) = 17 Then
                    
                    EL_CNT = (DateDiff("D", date_from, date_to) + 1)
                    'PL_CNT = PL_CNT + 1
                Else
                    EL_CNT = EL_CNT + 0.5
                End If
            Else
                EL_CNT = EL_CNT + 0.5
            End If
            
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
    GetemergencyLeave = EL_CNT
End Function
Private Sub cmdPrint_Click()
    Dim I                                                             As Integer
    Dim rsEmpInfo                                                     As ADODB.Recordset
    Dim rsLeave                                                       As ADODB.Recordset
    Dim EMPLNO                                                        As String
    Dim MEAL                                                          As Double
    Dim TRANSPO                                                       As Double
    Dim minlate                                                       As Integer
    Dim hourlate                                                      As Double
    Dim notimes                                                       As Integer

    If cboQuensina.Text = "1st Cut-Off" Then
        matt = "1"
    Else
        matt = "2"
    End If

    If Len(Dir(App.Path & "\DTR.html")) > 0 Then
        On Error Resume Next
        Kill (App.Path & "\DTR.html")
    End If
    
    Screen.MousePointer = 11
    Open App.Path & "\DTR.html" For Output As #1
    Set rsEmpInfo = gconDMIS.Execute("SELECT EMPNO, EMPLEVEL , LASTNAME, FIRSTNAME FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE <> 'I' AND EMPLEVEL <> 'M' ORDER BY LASTNAME")
    'Set rsEmpInfo = gconDMIS.Execute("SELECT EMPNO, EMPLEVEL , LASTNAME, FIRSTNAME FROM HRMS_EMPINFO ORDER BY LASTNAME")
    Print #1, "<html>"
    Print #1, "<head><title> DTR Summary Report</title>"
    Print #1, "<STYLE>"
    Print #1, "TH{TEXT-ALIGN:CENTER;FONT-FAMILY:ARIAL;FONT-SIZE:.8EM;FONT-WEIGHT:900;BACKGROUND-COLOR:#ccfdee;}"
    Print #1, "TD{TEXT-ALIGN:CENTER;FONT-FAMILY:ARIAL;FONT-SIZE:.8EM;}DIV{TEXT-ALIGN:LEFT;FONT-FAMILY:ARIAL;FONT-SIZE:14PX;FONT-WEIGHT:900;}"
    Print #1, "</STYLE>"
    Print #1, "</head>"
    Print #1, "<body>"

    Print #1, "<DIV>" & COMPANY_NAME & "</DIV>"
    Print #1, "<DIV>" & COMPANY_ADDRESS & "</DIV>"
    Print #1, "<DIV>WORK PERIOD: " & cboQuensina.Text & ", " & cboMOnth & " "; cboyear & "</DIV>"

    Print #1, "<table border=""1"" CELLPADDING=0 CELLSPACING=1 WIDTH=""100%"" STYLE=""border-collapse:collapse;"">"
    Print #1, "<TR>"
    Print #1, "<TH>&nbsp;</TH>"
    Print #1, "<TH>&nbsp;</TH>"
    Print #1, "<TH COLSPAN=3>TARDY/UNDER TIME</TH>"
    Print #1, "<TH COLSPAN=5>LEAVE WITH PAY</TH>"
    Print #1, "<TH>ABSENT</TH>"
    Print #1, "<TH COLSPAN=5>TOTAL O.T. HRS. CLAIM</TH>"
    Print #1, "<TH COLSPAN=4>REF.HOLIDAY/SPECIAL HOLIDAY</TH>"
    Print #1, "<TH>&nbsp;</TH>"
    Print #1, "<TH>&nbsp;</TH>"
    Print #1, "<TH>&nbsp;</TH>"
    Print #1, "</TR>"
    Print #1, "<TR>"
    Print #1, "<TH>&nbsp;</TH>"
    Print #1, "<TH>EMPLOYEE NAME</TH>"
    Print #1, "<TH WIDTH=""50PX"">#OF TIME</TH>"
    Print #1, "<TH WIDTH=""50PX"">#OF HRS.</TH>"
    Print #1, "<TH WIDTH=""50PX"">MINS.</TH>"
    Print #1, "<TH WIDTH=""50PX"">VL</TH>"
    Print #1, "<TH WIDTH=""50PX"">SL</TH>"
    'Print #1, "<TH WIDTH=""50PX"">PL/ML/EL</TH>"
    Print #1, "<TH WIDTH=""50PX"">PL</TH>"
    Print #1, "<TH WIDTH=""50PX"">ML</TH>"
    Print #1, "<TH WIDTH=""50PX"">EL</TH>"
    
    Print #1, "<TH WIDTH=""50PX"">LWOP</TH>"
    Print #1, "<TH WIDTH=""50PX"">REG</TH>"
    Print #1, "<TH WIDTH=""50PX"">ND/REG</TH>"
    Print #1, "<TH WIDTH=""50PX"">SUN</TH>"
    Print #1, "<TH WIDTH=""50PX"">OT/SUN</TH>"
    Print #1, "<TH WIDTH=""50PX"">ND/SUN</TH>"
    Print #1, "<TH WIDTH=""50PX"">RH</TH>"
    Print #1, "<TH WIDTH=""50PX"">RH/ND</TH>"
    Print #1, "<TH WIDTH=""50PX"">SH</TH>"
    Print #1, "<TH WIDTH=""50PX"">SH/ND</TH>"
    Print #1, "<TH WIDTH=""50PX"">MEAL</TH>"
    Print #1, "<TH WIDTH=""50PX"">TRANSPO</TH>"
    Print #1, "<TH>REMARKS</TH>"
    Print #1, "</tr>"

    prg.Max = rsEmpInfo.RecordCount
    prg.Value = 0
    While Not rsEmpInfo.EOF

        EMPLNO = Null2String(rsEmpInfo!EMPNO)

        I = I + 1
        Print #1, "<tr>"
        Print #1, "<td>"
        Print #1, I
        Print #1, "</td>"

        Print #1, "<td style=""text-align:left;"">"
        Print #1, rsEmpInfo!lastname & " ," & rsEmpInfo!FIRSTNAME
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "" & GetTardinessTimes(EMPLNO)
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; (GetTardiness(EMPLNO) - (GetTardiness(EMPLNO) Mod 60)) \ 60
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; GetTardiness(EMPLNO) Mod 60
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; GetVacationLeave(EMPLNO)
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; GetSickLeave(EMPLNO)
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; GetPaternityLeave(EMPLNO)
        Print #1, "</td>"


        Print #1, "<td>"
        Print #1, "&nbsp;"; GetMaternityLeave(EMPLNO)
        Print #1, "</td>"

        
        Print #1, "<td>"
        Print #1, "&nbsp;"; GetemergencyLeave(EMPLNO)
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; GetAbsencesTimes(EMPLNO)
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; HOUR_MIN(GetOT(EMPLNO, "001"))              'HOUR_MIN(OT_REG)
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; HOUR_MIN(GetOT(EMPLNO, "012"))              'HOUR_MIN(OT_ND_REG)
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; HOUR_MIN(GetOT(EMPLNO, "004"))              'HOUR_MIN(SUN)
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; HOUR_MIN(GetOT(EMPLNO, "009"))              'HOUR_MIN(OT_sun)
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; HOUR_MIN(GetOT(EMPLNO, "015"))              'HOUR_MIN(OT_ND_SUN)
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; HOUR_MIN(GetOT(EMPLNO, "002"))              'HOUR_MIN(HOL_REG)
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; HOUR_MIN(GetOT(EMPLNO, "013"))              'HOUR_MIN(HOL_REG_ND)
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; HOUR_MIN(GetOT(EMPLNO, "003"))              'HOUR_MIN(HOL_S)
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; HOUR_MIN(GetOT(EMPLNO, "014"))              'HOUR_MIN(HOL_SND)
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; GetMealTranspoAllowance(EMPLNO, "003")
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; GetMealTranspoAllowance(EMPLNO, "004")
        Print #1, "</td>"

        Print #1, "<td>"
        Print #1, "&nbsp;"; GetLeaveDates(EMPLNO, "VL") & GetLeaveDates(EMPLNO, "SL") & GetLeaveDates(EMPLNO, "EL") & GetLeaveDates(EMPLNO, "PL") & GetLeaveDates(EMPLNO, "ML") & GetAbsencesDates(EMPLNO) & GetOTDates(EMPLNO)
        Print #1, "</td>"
        Print #1, "</tr>"
        
        DoEvents
            prg.Value = prg.Value + 1
        DoEvents
        
        rsEmpInfo.MoveNext
    Wend
    Print #1, "</table>"

    Print #1, "<br/><br/><br/><table width=""100%"">"

    Print #1, "<TR><td>"
    Print #1, "Prepared By:"
    Print #1, "</td>"

    Print #1, "<td>"
    Print #1, "Checked By:"
    Print #1, "</td>"

    Print #1, "<td>"
    Print #1, "Noted By:"
    Print #1, "</td></TR>"

    Print #1, "<TR><td>"
    Print #1, "&nbsp;"; PREPARED_BY
    Print #1, "</td>"

    Print #1, "<td>"
    Print #1, "&nbsp;"; CHECKED_BY
    Print #1, "</td>"

    Print #1, "<td>"
    Print #1, "&nbsp;"; NOTED_BY
    Print #1, "</td></TR>"

    Print #1, "</table>"
    Print #1, "<span style=""width:100%;TEXT-ALIGN:right;FONT-FAMILY:ARIAL;FONT-SIZE:9PX;;"">Date Time Generated:" & Now & "<span>"
    Close #1
    Dim ie
    Set ie = CreateObject("InternetExplorer.Application")

    ie.Navigate2 App.Path & "\dtr.html"
    ie.Top = 0: ie.Left = 0
    ie.Width = Screen.Width - (Screen.TwipsPerPixelX) * 15
    ie.HEIGHT = Screen.HEIGHT - (Screen.TwipsPerPixelY) * 15
    ie.ToolBar = False
    ie.MenuBar = False
    ie.Visible = True
    
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
    cboQuensina.AddItem "1st Cut-Off"
    cboQuensina.AddItem "2nd Cut-Off"
    fillcbomonth cboMOnth
    'FillcboYear cboyear
    fillcombo_up cboyear
    If Day(LOGDATE) > PAYROLLCODE_TO1 And Day(LOGDATE) < PAYROLLCODE_TO2 Then
        cboQuensina.Text = "1st Cut-Off"
    Else
        cboQuensina.Text = "2nd Cut-Off"
    End If
    cboyear.Text = YEAR(LOGDATE)
    cboMOnth.Text = The_month(MONTH(LOGDATE))
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub


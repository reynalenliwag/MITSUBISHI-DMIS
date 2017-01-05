VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMSLoadApointmentToRO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Upload Appointment to Repair Order"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "FrmLoadApointmentToRO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   7290
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   30
      ScaleHeight     =   2865
      ScaleWidth      =   7215
      TabIndex        =   0
      Top             =   30
      Width           =   7245
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   855
         Left            =   6480
         MouseIcon       =   "FrmLoadApointmentToRO.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "FrmLoadApointmentToRO.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancel"
         Top             =   1980
         Width           =   735
      End
      Begin VB.TextBox txtAcct_No 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   270
         TabIndex        =   8
         Top             =   2220
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   5715
      End
      Begin VB.TextBox txtModel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   810
         Width           =   5715
      End
      Begin VB.TextBox txtPlanteNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   5490
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   1785
      End
      Begin VB.TextBox txtROno 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   5490
         MaxLength       =   8
         TabIndex        =   3
         Top             =   270
         Width           =   1755
      End
      Begin VB.TextBox txtKM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   1530
         TabIndex        =   2
         Top             =   1080
         Width           =   2205
      End
      Begin VB.ComboBox cboSA 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1350
         Width           =   3525
      End
      Begin VB.TextBox txtAppt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00004000&
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   2205
      End
      Begin VB.CommandButton cmdProces 
         Caption         =   "&Process Upload"
         Height          =   855
         Left            =   5760
         MouseIcon       =   "FrmLoadApointmentToRO.frx":1512
         MousePointer    =   99  'Custom
         Picture         =   "FrmLoadApointmentToRO.frx":1664
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Process Upload"
         Top             =   1980
         Width           =   735
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   285
         Left            =   0
         TabIndex        =   20
         Top             =   1650
         Width           =   7215
         _Version        =   655364
         _ExtentX        =   12726
         _ExtentY        =   503
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   285
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   7215
         _Version        =   655364
         _ExtentX        =   12726
         _ExtentY        =   503
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2BDB6&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Appointment No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   -15
         TabIndex        =   18
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2BDB6&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Customer Name "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   -15
         TabIndex        =   17
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2BDB6&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Model"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   0
         Left            =   -30
         TabIndex        =   16
         Top             =   810
         Width           =   1575
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2BDB6&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Plate No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   3720
         TabIndex        =   15
         Top             =   1080
         Width           =   1785
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2BDB6&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " New R/O No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   0
         Left            =   3720
         TabIndex        =   14
         Top             =   270
         Width           =   2205
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2BDB6&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " KM Reading"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   1
         Left            =   -30
         TabIndex        =   13
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2BDB6&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " R-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   1
         Left            =   5250
         TabIndex        =   12
         Top             =   270
         Width           =   2205
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00D2BDB6&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Service Advisor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   2
         Left            =   -30
         TabIndex        =   11
         Top             =   1350
         Width           =   3765
      End
   End
End
Attribute VB_Name = "frmCSMSLoadApointmentToRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim JobTotal                                           As Double
Dim JobComTotal                                        As Double
Dim JobSalesTotal                                      As Double
Dim JobWarTotal                                        As Double
Dim JobDiscTotal                                       As Double
Dim JobVatTotal                                        As Double

Dim PartsTotal                                         As Double
Dim PartsComTotal                                      As Double
Dim PartsSalesTotal                                    As Double
Dim PartsWarTotal                                      As Double
Dim PartsDiscTotal                                     As Double
Dim PartsVatTotal                                      As Double

Dim MatTotal                                           As Double
Dim MatComTotal                                        As Double
Dim MatSalesTotal                                      As Double
Dim MatWarTotal                                        As Double
Dim MatDiscTotal                                       As Double
Dim MatVatTotal                                        As Double
Dim COMTotal                                           As Double
Dim SALESTotal                                         As Double
Dim WARTotal                                           As Double
Dim VATTotal                                           As Double
Dim ROTotal                                            As Double
Dim Thedate                                            As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdProces_Click()

    On Error GoTo errorcode:
    If LTrim(RTrim(txtROno)) = "" Then
        ShowIsRequiredMsg "Repair Order no. cannot be Blank"
        txtROno.SetFocus
        Exit Sub
    End If
    If LTrim(RTrim(txtKM.Text)) = "" Then
        ShowIsRequiredMsg "KM Reading cannot be Blank"
        txtKM.SetFocus
        Exit Sub
    End If

    Dim VTXTRONO                                       As String
    VTXTRONO = "R-" & Format(txtROno, "00000000")

    Thedate = Format(Now, "MM/dd/yyyy")

    Dim rsREPOR                                        As New ADODB.Recordset
    Set rsREPOR = gconDMIS.Execute("SELECT REP_OR,NIYM FROM CSMS_REPOR WHERE REP_OR = '" & VTXTRONO & "' and TRANSTYPE = 'R'")
    If Not (rsREPOR.BOF And rsREPOR.EOF) Then
        MsgBox "Repair Order No. " & VTXTRONO & ": " & Null2String(rsREPOR!NIYM) & " already Exist", vbExclamation, "CSMS"
        txtROno.SetFocus
        Exit Sub
    End If
    Set rsREPOR = Nothing

    If MsgBox("Upload Appointment no. " & txtAppt & " to Repair Order no. " & VTXTRONO & "", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub

    SQL_STATEMENT = "update CSMS_Repor set" & _
        " REP_OR = '" & VTXTRONO & _
        "', DTE_RECD = '" & Thedate & _
        "', transtype = '" & "R" & _
        "', KM_RDG = " & N2Str2Null(txtKM) & _
        ", RECD_BY = " & N2Str2Null(GetSaCode(cboSA)) & _
        " where ApptNo = '" & txtAppt & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("UP", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtAppt), "ApptNo", "CSMS_REPOR"), "R", "RO NO: " & VTXTRONO, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    SQL_STATEMENT = "update CSMS_Ro_Det set" & _
        " REP_OR = '" & VTXTRONO & "'," & _
        " transtype = 'R'" & _
        " where ApptNo = '" & txtAppt & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("UD", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtAppt), "ApptNo", "CSMS_REPOR"), "R", "RO NO: " & VTXTRONO & " - JOBS", "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    SQL_STATEMENT = "update CSMS_RepairOrder set" & _
        " RO_No = '" & VTXTRONO & _
        "',appointmentdate='" & Thedate & "'," & _
        " transtype = 'R'" & _
        ", WRITER = " & N2Str2Null(cboSA) & _
        " where ApptNo = '" & txtAppt & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("UD", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtAppt), "ApptNo", "CSMS_REPOR"), "R", "RO NO: " & VTXTRONO & " - SERVICE COUNTER", "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    SQL_STATEMENT = "update CSMS_PMS_Job_Det set" & _
        " REP_OR = '" & VTXTRONO & "'," & _
        " transtype = 'R'" & _
        " where ApptNo = '" & txtAppt & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("UD", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtAppt), "ApptNo", "CSMS_REPOR"), "R", "RO NO: " & VTXTRONO & " - PMS JOBS", "", "")
    'NEW LOG AUDIT-----------------------------------------------------


    Dim VTXTREP_OR                                     As String

    VTXTREP_OR = N2Str2Null(txtROno)
    SQL_STATEMENT = "update CSMS_Appointment set " & _
        " status = 'Served' " & _
        ",KM_RDG = " & N2Str2Null(txtKM) & _
        " where ApptNo = '" & txtAppt & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("UD", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtAppt), "ApptNo", "CSMS_REPOR"), "R", "RO NO: " & txtROno & " - APPOINTMENT", "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Dim rsRO_DET                                       As New ADODB.Recordset
    TOTJOBAMT = 0: TOTJOBDISC = 0: TOTJOBDISCVAL = 0: TOTJOBTAX = 0
    JobComTotal = 0: JobSalesTotal = 0: JobWarTotal = 0
    Set rsRO_DET = gconDMIS.Execute("select discount_2,det_amt,wcode,disval,taxval from CSMS_RO_Det where rep_or = " & VTXTREP_OR & " and livil = '1' order by LINE_NO asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Screen.MousePointer = 11
        rsRO_DET.MoveFirst
        Do While Not rsRO_DET.EOF
            If Null2String(rsRO_DET!wCode) = "C" Then
                JobComTotal = JobComTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then JobSalesTotal = JobSalesTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then JobWarTotal = JobWarTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            Else
                TOTJOBAMT = TOTJOBAMT + N2Str2Zero(rsRO_DET!DET_AMT)
                TOTJOBDISC = TOTJOBDISC + N2Str2Zero(rsRO_DET!Discount_2)
                TOTJOBDISCVAL = TOTJOBDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTJOBTAX = TOTJOBTAX + N2Str2Zero(rsRO_DET!TAXVAL)
            End If
            rsRO_DET.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRO_DET = Nothing
    TOTJOBAMT = Round(TOTJOBAMT, 2): TOTJOBDISC = Round(TOTJOBDISC, 2): TOTJOBDISCVAL = Round(TOTJOBDISCVAL, 2): TOTJOBTAX = Round(TOTJOBTAX, 2)


    TOTPARTSAMT = 0: TOTPARTSDISC = 0: TOTPARTSDISCVAL = 0: TOTPARTSTAX = 0
    PartsComTotal = 0: PartsSalesTotal = 0: PartsWarTotal = 0

    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where rep_or = " & VTXTREP_OR & " and livil = '2' order by LINE_NO asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        rsRO_DET.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsRO_DET.EOF
            If Null2String(rsRO_DET!wCode) = "C" Then
                PartsComTotal = PartsComTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then PartsSalesTotal = PartsSalesTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then PartsWarTotal = PartsWarTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            Else
                TOTPARTSAMT = TOTPARTSAMT + N2Str2Zero(rsRO_DET!DET_AMT)
                TOTPARTSDISC = TOTPARTSDISC + N2Str2Zero(rsRO_DET!Discount_2)
                TOTPARTSDISCVAL = TOTPARTSDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTPARTSTAX = TOTPARTSTAX + N2Str2Zero(rsRO_DET!TAXVAL)
            End If
            rsRO_DET.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRO_DET = Nothing
    TOTPARTSAMT = Round(TOTPARTSAMT, 2): TOTPARTSDISC = Round(TOTPARTSDISC, 2): TOTPARTSDISCVAL = Round(TOTPARTSDISCVAL, 2): TOTPARTSTAX = Round(TOTPARTSTAX, 2)

    TOTMATAMT = 0: TOTMATDISC = 0: TOTMATDISCVAL = 0: TOTMATTAX = 0
    MatComTotal = 0: MatSalesTotal = 0: MatWarTotal = 0

    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where rep_or = " & VTXTREP_OR & " and livil = '3' order by LINE_NO asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Screen.MousePointer = 11
        rsRO_DET.MoveFirst
        Do While Not rsRO_DET.EOF
            If Null2String(rsRO_DET!wCode) = "C" Then
                MatComTotal = MatComTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then MatSalesTotal = MatSalesTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then MatWarTotal = MatWarTotal + N2Str2Zero(rsRO_DET!DET_AMT)
            Else
                TOTMATAMT = TOTMATAMT + N2Str2Zero(rsRO_DET!DET_AMT)
                TOTMATDISC = TOTMATDISC + N2Str2Zero(rsRO_DET!Discount_2)
                TOTMATDISCVAL = TOTMATDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTMATTAX = TOTMATTAX + N2Str2Zero(rsRO_DET!TAXVAL)
            End If
            rsRO_DET.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRO_DET = Nothing
    TOTMATAMT = Round(TOTMATAMT, 2): TOTMATDISC = Round(TOTMATDISC, 2): TOTMATDISCVAL = Round(TOTMATDISCVAL, 2): TOTMATTAX = Round(TOTMATTAX, 2)

    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT
    SQL_STATEMENT = "update CSMS_RepOr set" & _
        " labor = " & Round(TOTJOBAMT - TOTJOBTAX, 2) & "," & _
        " l_amtvalue = " & Round(TOTJOBAMT, 2) & "," & _
        " l_disc = " & Round(TOTJOBDISCVAL, 2) & "," & _
        " l_disc2 = " & Round(TOTJOBDISC * (VAT_RATE / 100), 2) & "," & _
        " l_taxval = " & Round(TOTJOBTAX, 2) & "," & _
        " l_discount = " & Round(TOTJOBDISC, 2) & "," & _
        " amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC, 2) & "," & _
        " rovat = " & Round(TOTJOBTAX + TOTPARTSTAX + TOTMATTAX, 2) & "," & _
        " wl_amt = " & 0 & "," & _
        " ro_amount = " & Round(ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC, 2) & _
        " where REP_OR = " & VTXTREP_OR
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtAppt), "ApptNo", "CSMS_REPOR"), "R", "RO NO: " & txtROno & " - UPLOAD APPOINTMENT", "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    MessagePop InfoFriend, "Appointment Information Updated", "Appointment Information Sucessfully Uploaded!", 1000
    cmdCancel.Value = True

    If MODULENAME = "CSMS" Then
        If FROM_APPOINTMENT = "DASH" Then Exit Sub

        Call frmCSMS_ServiceCounter.Click_ScheduleGrid
    End If
    
    Exit Sub
errorcode:
    MsgBox Err.Description
    Exit Sub
    
End Sub

Function GetSaCode(xName As String) As String
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT CODE FROM CSMS_VW_EMPNO WHERE NAYM = " & N2Str2Null(xName) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetSaCode = Null2String(RSTMP!Code)
    End If
    Set RSTMP = Nothing
End Function

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    Call FillSA
End Sub

Private Sub lblLogLoan_Click()

End Sub

Private Sub txtKM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtROno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Sub FillSA()
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT (LASTNAME + ', ' + FIRSTNAME + ' ' + LEFT(MIDDLENAME,1) + '.') AS NAYM " & _
        " FROM HRMS_EMPINFO WHERE IS_SERVICE_ADVISER = '1' " & _
        " AND ACTIVEINACTIVE = 'A' " & _
        " AND RESIGNED IS NULL")
    cboSA.Clear
    If Not (RSTMP.EOF And RSTMP.BOF) Then
        RSTMP.MoveFirst
        
        Do While Not RSTMP.EOF
            cboSA.AddItem Null2String(RSTMP!NAYM)
            RSTMP.MoveNext
        Loop
        cboSA.ListIndex = 0
    End If
    Set RSTMP = Nothing
End Sub

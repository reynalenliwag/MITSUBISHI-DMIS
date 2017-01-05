VERSION 5.00
Begin VB.Form frmCSMSLoadApointmentToRO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Upload Appointment to Repair Order"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "FrmLoadApointmentToRO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   7140
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   105
      Left            =   -150
      TabIndex        =   10
      Top             =   1740
      Width           =   7455
   End
   Begin VB.TextBox txtROno 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   5160
      TabIndex        =   9
      Top             =   240
      Width           =   1875
   End
   Begin VB.TextBox txtPlanteNo 
      BackColor       =   &H8000000F&
      Height          =   405
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   1875
   End
   Begin VB.TextBox txtModel 
      BackColor       =   &H8000000F&
      Height          =   405
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   1755
   End
   Begin VB.TextBox txtCustomer 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   5385
   End
   Begin VB.TextBox txtAppt 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   1725
   End
   Begin VB.TextBox txtAcct_No 
      Height          =   375
      Left            =   1950
      TabIndex        =   11
      Top             =   720
      Width           =   525
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
      Height          =   735
      Left            =   6300
      MouseIcon       =   "FrmLoadApointmentToRO.frx":014A
      MousePointer    =   99  'Custom
      Picture         =   "FrmLoadApointmentToRO.frx":029C
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cancel"
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdProces 
      Caption         =   "&Process Upload"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5580
      MouseIcon       =   "FrmLoadApointmentToRO.frx":05DA
      MousePointer    =   99  'Custom
      Picture         =   "FrmLoadApointmentToRO.frx":072C
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Process Upload"
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "New R/O No."
      Height          =   315
      Left            =   3960
      TabIndex        =   8
      Top             =   330
      Width           =   1545
   End
   Begin VB.Label Label4 
      Caption         =   "Plate No."
      Height          =   315
      Left            =   4320
      TabIndex        =   6
      Top             =   1290
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Model"
      Height          =   315
      Left            =   1020
      TabIndex        =   4
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label Label2 
      Caption         =   "Customer Name "
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   780
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Appointment No."
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   330
      Width           =   1545
   End
End
Attribute VB_Name = "frmCSMSLoadApointmentToRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim JobTotal                            As Double
Dim JobComTotal                         As Double
Dim JobSalesTotal                       As Double
Dim JobWarTotal                         As Double
Dim JobDiscTotal                        As Double
Dim JobVatTotal                         As Double

Dim PartsTotal                          As Double
Dim PartsComTotal                       As Double
Dim PartsSalesTotal                     As Double
Dim PartsWarTotal                       As Double
Dim PartsDiscTotal                      As Double
Dim PartsVatTotal                       As Double

Dim MatTotal                            As Double
Dim MatComTotal                         As Double
Dim MatSalesTotal                       As Double
Dim MatWarTotal                         As Double
Dim MatDiscTotal                        As Double
Dim MatVatTotal                         As Double
Dim COMTotal                            As Double
Dim SALESTotal                          As Double
Dim WARTotal                            As Double
Dim VATTotal                            As Double
Dim ROTotal                             As Double
Dim Thedate                             As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdProces_Click()
        Dim Rep_Or2, rep_or3                As String
        Dim k                               As Integer
        If Left(txtROno.Text, 2) = "R-" Then
            txtROno.Text = "R-" & Format(NumericVal(Right(txtROno.Text, Len(txtROno.Text) - 2)), "00000000")
        Else
            txtROno.Text = "R-" & Format(NumericVal(Right(txtROno.Text, Len(txtROno.Text))), "00000000")
        End If
        
            Dim rsReporDup                  As ADODB.Recordset
            Set rsReporDup = New ADODB.Recordset
            Set rsReporDup = gconDMIS.Execute("select RO_NO from CSMS_REPAIRORDER where RO_NO= " & N2Str2Null(txtROno.Text))
            If Not rsReporDup.EOF And Not rsReporDup.BOF Then
                MsgSpeechBox "Warning: Repair Order Number Already Exist!"
                On Error Resume Next
                txtROno.SetFocus
            End If
            Set rsReporDup = Nothing
    
    
    
    Thedate = Format(Now, "MM/dd/yyyy")

    gconDMIS.Execute "update CSMS_Repor set" & _
                   " REP_OR = '" & txtROno & "',DTE_RECD='" & Thedate & "'," & _
                   " transtype = 'R'" & _
                   " where ApptNo = '" & txtAppt & "'"




    gconDMIS.Execute "update CSMS_Ro_Det set" & _
                   " REP_OR = '" & txtROno & "'," & _
                   " transtype = 'R'" & _
                   " where ApptNo = '" & txtAppt & "'"

    gconDMIS.Execute "update CSMS_RepairOrder set" & _
                   " RO_No = '" & txtROno & "',appointmentdate='" & Thedate & "'," & _
                   " transtype = 'R'" & _
                   " where ApptNo = '" & txtAppt & "'"

    gconDMIS.Execute "update CSMS_PMS_Job_Det set" & _
                   " REP_OR = '" & txtROno & "'," & _
                   " transtype = 'R'" & _
                   " where ApptNo = '" & txtAppt & "'"
    Dim VtxtROno                      As String
    VtxtROno = N2Str2Null(txtROno)
    gconDMIS.Execute "update CSMS_Appointment set status='Served' where ApptNo = '" & txtAppt & "'"
    Dim rsRO_DET                        As ADODB.Recordset
    TOTJOBAMT = 0: TOTJOBDISC = 0: TOTJOBDISCVAL = 0: TOTJOBTAX = 0
    JobComTotal = 0: JobSalesTotal = 0: JobWarTotal = 0
    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select discount_2,det_amt,wcode,disval,taxval from CSMS_RO_Det where rep_or = " & VtxtROno & " and livil = '1' order by LINE_NO asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Screen.MousePointer = 11
        rsRO_DET.MoveFirst
        Do While Not rsRO_DET.EOF
            If Null2String(rsRO_DET!wCode) = "C" Then
                JobComTotal = JobComTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then JobSalesTotal = JobSalesTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then JobWarTotal = JobWarTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            Else
                TOTJOBAMT = TOTJOBAMT + N2Str2Zero(rsRO_DET!Det_AMT)
                TOTJOBDISC = TOTJOBDISC + N2Str2Zero(rsRO_DET!discount_2)
                TOTJOBDISCVAL = TOTJOBDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTJOBTAX = TOTJOBTAX + N2Str2Zero(rsRO_DET!taxval)
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
    Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where rep_or = " & VtxtROno & " and livil = '2' order by LINE_NO asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        rsRO_DET.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsRO_DET.EOF
            If Null2String(rsRO_DET!wCode) = "C" Then
                PartsComTotal = PartsComTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then PartsSalesTotal = PartsSalesTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then PartsWarTotal = PartsWarTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            Else
                TOTPARTSAMT = TOTPARTSAMT + N2Str2Zero(rsRO_DET!Det_AMT)
                TOTPARTSDISC = TOTPARTSDISC + N2Str2Zero(rsRO_DET!discount_2)
                TOTPARTSDISCVAL = TOTPARTSDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTPARTSTAX = TOTPARTSTAX + N2Str2Zero(rsRO_DET!taxval)
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
    Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where rep_or = " & VtxtROno & " and livil = '3' order by LINE_NO asc")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Screen.MousePointer = 11
        rsRO_DET.MoveFirst
        Do While Not rsRO_DET.EOF
            If Null2String(rsRO_DET!wCode) = "C" Then
                MatComTotal = MatComTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "S" Then MatSalesTotal = MatSalesTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            ElseIf Null2String(rsRO_DET!wCode) = "W" Then MatWarTotal = MatWarTotal + N2Str2Zero(rsRO_DET!Det_AMT)
            Else
                TOTMATAMT = TOTMATAMT + N2Str2Zero(rsRO_DET!Det_AMT)
                TOTMATDISC = TOTMATDISC + N2Str2Zero(rsRO_DET!discount_2)
                TOTMATDISCVAL = TOTMATDISCVAL + N2Str2Zero(rsRO_DET!disval)
                TOTMATTAX = TOTMATTAX + N2Str2Zero(rsRO_DET!taxval)
            End If
            rsRO_DET.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set rsRO_DET = Nothing
    TOTMATAMT = Round(TOTMATAMT, 2): TOTMATDISC = Round(TOTMATDISC, 2): TOTMATDISCVAL = Round(TOTMATDISCVAL, 2): TOTMATTAX = Round(TOTMATTAX, 2)

    ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT
    gconDMIS.Execute "update CSMS_RepOr set" & _
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
                   " where REP_OR = " & VtxtROno
    cmdCancel.Value = True
    LogAudit "P", "JOB ADDED", " APPOINTMENT:" & txtAppt & " RO:" & txtROno
    frmCSMSAppointment.ViewAppointmentGrid
End Sub


Private Sub txtROno_Change()
    If txtROno = "" Then
    Else
    End If
End Sub

Private Sub txtROno_Validate(Cancel As Boolean)
    Dim Rep_Or2, rep_or3                As String
    Dim k                               As Integer
    If Left(txtROno.Text, 2) = "R-" Then
        txtROno.Text = "R-" & Format(NumericVal(Right(txtROno.Text, Len(txtROno.Text) - 2)), "00000000")
    Else
        txtROno.Text = "R-" & Format(NumericVal(Right(txtROno.Text, Len(txtROno.Text))), "00000000")
    End If
End Sub

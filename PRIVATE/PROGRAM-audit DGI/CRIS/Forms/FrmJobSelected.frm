VERSION 5.00
Begin VB.Form frmCSMSJobSelected 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selected Job"
   ClientHeight    =   5865
   ClientLeft      =   75
   ClientTop       =   435
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "FrmJobSelected.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   7800
   StartUpPosition =   1  'CenterOwner
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
      Left            =   6930
      MouseIcon       =   "FrmJobSelected.frx":01CA
      MousePointer    =   99  'Custom
      Picture         =   "FrmJobSelected.frx":031C
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Cancel"
      Top             =   5010
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
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
      Left            =   6210
      MouseIcon       =   "FrmJobSelected.frx":065A
      MousePointer    =   99  'Custom
      Picture         =   "FrmJobSelected.frx":07AC
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Save Entry"
      Top             =   5010
      Width           =   735
   End
   Begin VB.TextBox txtROno 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   6300
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   120
      Width           =   1365
   End
   Begin VB.CommandButton cmdunloackSR 
      Height          =   345
      Left            =   6420
      Picture         =   "FrmJobSelected.frx":0A47
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2610
      Width           =   405
   End
   Begin VB.TextBox txtstdrate 
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   5580
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2610
      Width           =   765
   End
   Begin VB.Frame Frame1 
      Caption         =   "Job Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   90
      TabIndex        =   2
      Top             =   630
      Width           =   7575
      Begin VB.TextBox txtDetCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1710
         TabIndex        =   8
         Top             =   1560
         Width           =   1245
      End
      Begin VB.TextBox txtOPCODE 
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1710
         TabIndex        =   27
         Top             =   810
         Width           =   1725
      End
      Begin VB.ComboBox cboJobChargeTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   5490
         Sorted          =   -1  'True
         TabIndex        =   20
         Text            =   "cboJobChargeTo"
         Top             =   2400
         Width           =   795
      End
      Begin VB.TextBox txtJobDiscount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "0"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtNote 
         BackColor       =   &H00FFFFFF&
         Height          =   1230
         Left            =   1710
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "FrmJobSelected.frx":0FD1
         Top             =   2880
         Width           =   5655
      End
      Begin VB.CommandButton cmdunloackFR 
         Height          =   375
         Left            =   3090
         Picture         =   "FrmJobSelected.frx":0FD7
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1980
         Width           =   405
      End
      Begin VB.TextBox txtflatrate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1980
         Width           =   1245
      End
      Begin VB.TextBox txtJobDesc 
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   1080
         Left            =   3480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   810
         Width           =   3855
      End
      Begin VB.TextBox txtJobCat 
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   390
         Width           =   5625
      End
      Begin VB.TextBox txtCheckMe 
         BackColor       =   &H000000FF&
         Height          =   375
         Left            =   3660
         TabIndex        =   26
         Top             =   1320
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox txtjCode 
         BackColor       =   &H8000000F&
         Height          =   360
         Left            =   3720
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label labDetCost 
         Alignment       =   1  'Right Justify
         Caption         =   "Job Cost"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   600
         TabIndex        =   30
         Top             =   1590
         Width           =   1005
      End
      Begin VB.Label Label27 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Charge To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   4440
         TabIndex        =   23
         Top             =   2460
         Width           =   1305
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   480
         TabIndex        =   22
         Top             =   2430
         Width           =   1095
      End
      Begin VB.Label Label55 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2310
         TabIndex        =   21
         Top             =   2460
         Width           =   225
      End
      Begin VB.Label Label1 
         Caption         =   "Note/ Suggested Jobs"
         Height          =   615
         Index           =   7
         Left            =   210
         TabIndex        =   16
         Top             =   2910
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "Standard Time Rate"
         Height          =   315
         Index           =   6
         Left            =   3660
         TabIndex        =   14
         Top             =   2040
         Width           =   2445
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Flat Rate"
         Height          =   315
         Index           =   4
         Left            =   600
         TabIndex        =   7
         Top             =   2010
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Job Description"
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   870
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Job Category"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   420
         Width           =   1305
      End
   End
   Begin VB.TextBox txtCustomer 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.TextBox txtSaveorEdit 
      BackColor       =   &H8000000F&
      Height          =   360
      Left            =   2130
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   120
      Width           =   1305
   End
   Begin VB.TextBox txtAppt 
      Height          =   345
      Left            =   2400
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Flat Rate"
      Height          =   315
      Index           =   5
      Left            =   3840
      TabIndex        =   11
      Top             =   3150
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Customer "
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   885
   End
End
Attribute VB_Name = "frmCSMSJobSelected"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUpload                            As ADODB.Recordset
Dim AddorEdit                           As String

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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Screen.MousePointer = 11
    If txtCheckMe.Text = "ro" Then
        With frmCSMSNewAppointment.lblJob4Service
            .Sorted = False
            .ListItems.Add , , txtjCode
            .ListItems(.ListItems.Count).ListSubItems.Add 1, , "GJ"
            .ListItems(.ListItems.Count).ListSubItems.Add 2, , txtjobdesc.Text
            .ListItems(.ListItems.Count).ListSubItems.Add 3, , txtflatrate.Text
            .ListItems(.ListItems.Count).ListSubItems.Add 4, , txtstdrate.Text
            .ListItems(.ListItems.Count).ListSubItems.Add 5, , txtJobDiscount.Text
            .ListItems(.ListItems.Count).ListSubItems.Add 6, , cboJobChargeTo.Text
            .ListItems(.ListItems.Count).ListSubItems.Add 7, , txtnote.Text
        End With
    End If
    If txtCheckMe.Text = "app" Then
        With frmCSMSNewAppointment.lblJob4Service
            .Sorted = False
            .ListItems.Add , , txtjCode.Text
            .ListItems(.ListItems.Count).ListSubItems.Add 1, , txtjobdesc.Text
            .ListItems(.ListItems.Count).ListSubItems.Add 2, , txtflatrate.Text
            .ListItems(.ListItems.Count).ListSubItems.Add 3, , txtstdrate.Text
            .ListItems(.ListItems.Count).ListSubItems.Add 4, , txtJobDiscount.Text
            .ListItems(.ListItems.Count).ListSubItems.Add 5, , cboJobChargeTo.Text
            .ListItems(.ListItems.Count).ListSubItems.Add 6, , txtnote.Text
        End With
    End If

    If txtCheckMe.Text = "main" Then
        frmCSMSNewAppointment.lblJob4Service.ListItems.Clear
        Call SaveRePairOrder

        Call checkIFFinish
    End If

    If txtCheckMe.Text = "est" Then
        'With frmCSMSEstimateEntry
        '    .txtJobLineNo.Text = Format(ESTIKCNT + 1, "00")
        '    .txtJobPostCode.Text = ""
        '    .cboJobChargeTo.Clear
        '    .cboJobChargeTo.AddItem ""
        '    .cboJobChargeTo.AddItem "W"
        '    .cboJobChargeTo.AddItem "S"
        '    .cboJobChargeTo.AddItem "C"
        '    .txtJobDiscount.Text = "0"
        '    .txtJobDetail.Text = ""
        '    .cboJcode.Text = txtjCode.Text
        '    .cboJobCode.Text = txtJobDesc.Text
        '    .txtJobRate.Text = Val(txtflatrate.Text) * Val(txtstdrate.Text)
        '    .cmdJobSave.Value = True
        'End With
    End If

    Unload Me
    frmCSMSReqJobs.cmdClose.Value = True
    Screen.MousePointer = 0


    Exit Sub
End Sub

Sub SaveRePairOrder()

    Dim JOBREP_OR, JOBLEVEL, JOBLINE_NO, JOBDETCDE, VLastUpdateTime As String
    Dim JOBDETDSC, JOBDETUNT, VLastUpdate, Vusercode As String
    Dim JOBDETVOL, JOBDETPRC, JOBDETAMT As Double
    Dim JOBCODE, JOBWCODE, xApptNo      As String
    Dim JOBTAXRATE, JOBDISCRATE         As Double
    Dim JOBTAXVAL, JOBDISVAL            As Double
    Dim JOBPOCODE, JOBRep_Or2, JOBDETAIL As String
    Dim JOBDET_AMT, JOBDIS_VAL, JOBDISCOUNT_2, xFLATRATE As Double
    Dim JOBREMARKS                      As String
    Dim JOBTECHNICIAN                   As String
    Dim JOBDET_HRS                      As String
    Dim TheDone                         As String
    Dim JOBDETCOST                      As Double

    JOBDISVAL = 0: JOBTAXVAL = 0: JOBDETAMT = 0
    JOBDIS_VAL = 0: JOBDISCOUNT_2 = 0: JOBDISCRATE = 0

    JOBREP_OR = N2Str2Null(txtROno.Text)
    JOBLEVEL = "'1'"
    JOBLINE_NO = GetJobLineno(txtROno)
    JOBDETCDE = N2Str2Null(txtjCode.Text)
    JOBDETDSC = N2Str2Null(Mid(txtjobdesc.Text, 1, 500))
    JOBDETUNT = "NULL"
    JOBDETVOL = NumericVal(0)
    JOBDET_HRS = NumericVal(txtstdrate.Text)
    xFLATRATE = NumericVal(txtflatrate)
    JOBDETCOST = NumericVal(txtDetCost.Text)
    JOBDETPRC = NumericVal(txtflatrate.Text) * JOBDET_HRS
    JOBCODE = "NULL"
    JOBWCODE = N2Str2Null(cboJobChargeTo.Text)
    JOBTAXRATE = (VAT_RATE / 100)
    JOBDISCRATE = NumericVal(txtJobDiscount.Text) / 100
    JOBDETAMT = JOBDETPRC / ConvertToBIRDecimalFormat(VAT_RATE)
    JOBDISVAL = (JOBDETPRC * JOBDISCRATE) - ((JOBDETPRC * JOBDISCRATE) * JOBTAXRATE)
    JOBPOCODE = "NULL"
    JOBRep_Or2 = "NULL"
    JOBDETAIL = N2Str2Null(CheckChar(txtnote.Text))
    JOBDET_AMT = JOBDETPRC
    JOBDIS_VAL = JOBDISVAL * ConvertToBIRDecimalFormat(VAT_RATE)
    JOBDISCOUNT_2 = JOBDET_AMT * JOBDISCRATE
    JOBREMARKS = N2Str2Null(CheckChar(txtnote.Text))
    JOBTECHNICIAN = "NULL"
    JOBTAXVAL = Round(((JOBDETAMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"
    VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"
    xApptNo = N2Str2Null(txtAppt)
    TheDone = "'N'"
    
    gconDMIS.Execute "insert into CSMS_RO_Det " & _
                     "(ApptNo,FLATRATE,rep_or,livil,LINE_NO,detcde,detdsc,technician,det_hrs,detunt,detvol,detcost,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,USERCDE,SAVEDATE,SAVETIME,Done)" & _
                   " values (" & xApptNo & "," & xFLATRATE & "," & JOBREP_OR & ", " & JOBLEVEL & ", " & JOBLINE_NO & "," & _
                   " " & JOBDETCDE & "," & JOBDETDSC & "," & JOBTECHNICIAN & "," & JOBDET_HRS & "," & _
                   " " & JOBDETUNT & ", " & JOBDETVOL & "," & _
                   " " & JOBDETCOST & "," & JOBDETPRC & ", " & JOBDETAMT & ", " & JOBCODE & _
                     ", " & JOBWCODE & ", " & (JOBTAXRATE * 100) & ", " & (JOBDISCRATE * 100) & _
                     ", " & JOBTAXVAL & ", " & JOBDISVAL & ", " & JOBPOCODE & _
                     ", " & JOBRep_Or2 & ", " & JOBDETAIL & ", " & JOBDET_AMT & _
                     ", " & JOBDIS_VAL & ", " & JOBDISCOUNT_2 & _
                     ", " & Vusercode & _
                     ", " & VLastUpdate & _
                     ", " & VLastUpdateTime & "," & TheDone & ")"
    Dim rsRO_DET                        As ADODB.Recordset
    TOTJOBAMT = 0: TOTJOBDISC = 0: TOTJOBDISCVAL = 0: TOTJOBTAX = 0
    JobComTotal = 0: JobSalesTotal = 0: JobWarTotal = 0
    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("select discount_2,det_amt,wcode,disval,taxval from CSMS_RO_Det where rep_or = " & JOBREP_OR & " and livil = '1' order by LINE_NO asc")
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
    Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where rep_or = " & JOBREP_OR & " and livil = '2' order by LINE_NO asc")
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
    Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where rep_or = " & JOBREP_OR & " and livil = '3' order by LINE_NO asc")
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
                   " where REP_OR = " & JOBREP_OR

LogAudit "A", "JOB ADDED", "JCODE:" & txtjCode & " RO:" & txtROno
End Sub

Function GetJobLineno(XXX As String)
    Dim rsGetLN                         As ADODB.Recordset
    Set rsGetLN = New ADODB.Recordset
    Set rsGetLN = gconDMIS.Execute("Select [LINE_NO],REP_OR from CSMS_Ro_Det where [REP_OR] = '" & XXX & "' order by [LINE_NO] desc")
    If Not rsGetLN.EOF And Not rsGetLN.BOF Then
        GetJobLineno = Null2String(Format(Val(rsGetLN!LINE_NO) + 1, "00"))
    Else
        GetJobLineno = N2Str2Null(Format(1, "00"))
    End If
    Set rsGetLN = Nothing
End Function

Private Sub cmdunloackFR_Click()
    If txtflatrate.Locked = True Then
        txtflatrate.Locked = False
        txtflatrate.BackColor = &HFFFFFF
        txtflatrate.ForeColor = &H0&
    Else
        txtflatrate.Locked = True
        txtflatrate.BackColor = &HFF8080
        txtflatrate.ForeColor = &HFFFFFF
    End If
End Sub

Private Sub cmdunloackSR_Click()
    If txtstdrate.Locked = True Then
        txtstdrate.Locked = False
        txtstdrate.BackColor = &HFFFFFF
        txtstdrate.ForeColor = &H0&
    Else
        txtstdrate.Locked = True
        txtstdrate.BackColor = &HFF8080
        txtstdrate.ForeColor = &HFFFFFF
    End If
End Sub

Private Sub Command1_Click()
    Call checkIFFinish
End Sub

Private Sub txtJobDesc_Change()
    txtnote = txtjobdesc.Text
End Sub

Sub checkIFFinish()
    Dim SQL                             As String
    Dim RS                              As New ADODB.Recordset
    Dim theRo                           As String

    theRo = Trim(txtROno.Text)

    SQL = "SELECT Status FROM CSMS_repairOrder Where Ro_no='" & theRo & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    With RS

        If Not .EOF And Not .BOF Then

            If StrComp(Trim(!STATUS), "Finish Job") = 0 Then
                gconDMIS.Execute "UPDATE CSMS_repairOrder SET dateFinish=NULL,jstatus='W',status='Working' where Ro_no='" & theRo & "'"
                MsgBox "New Job has Been Added.Pls be advice to assigned a Technician", vbInformation, "Confirm"
            End If

        End If

    End With

    Set RS = Nothing
End Sub

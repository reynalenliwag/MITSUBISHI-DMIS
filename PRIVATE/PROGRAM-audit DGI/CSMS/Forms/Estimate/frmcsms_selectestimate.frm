VERSION 5.00
Begin VB.Form frmcsms_selectestimate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Estimate"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtjCode 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2880
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   5160
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Frame Frame1 
      Caption         =   "JOB INFORMATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4275
      Left            =   0
      TabIndex        =   7
      Top             =   510
      Width           =   7575
      Begin VB.TextBox txtJobCat 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   390
         Width           =   5625
      End
      Begin VB.TextBox txtJobDesc 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1080
         Left            =   3480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   810
         Width           =   3855
      End
      Begin VB.TextBox txtflatrate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1980
         Width           =   1245
      End
      Begin VB.CommandButton cmdunloackFR 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2970
         Picture         =   "frmcsms_selectestimate.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1980
         Width           =   405
      End
      Begin VB.TextBox txtNote 
         BackColor       =   &H00FFFFFF&
         Height          =   1230
         Left            =   1710
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "frmcsms_selectestimate.frx":058A
         Top             =   2880
         Width           =   5655
      End
      Begin VB.TextBox txtJobDiscount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
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
         TabIndex        =   14
         Text            =   "0"
         Top             =   2400
         Width           =   495
      End
      Begin VB.ComboBox cboJobChargeTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
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
         TabIndex        =   13
         Text            =   "cboJobChargeTo"
         Top             =   2430
         Width           =   795
      End
      Begin VB.TextBox txtOPCODE 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1710
         TabIndex        =   12
         Top             =   810
         Width           =   1725
      End
      Begin VB.TextBox txtDetCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1710
         TabIndex        =   11
         Top             =   1560
         Width           =   1245
      End
      Begin VB.ComboBox cboBP_TYPE 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmcsms_selectestimate.frx":0590
         Left            =   2700
         List            =   "frmcsms_selectestimate.frx":059A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2430
         Width           =   1605
      End
      Begin VB.CommandButton cmdunloackSR 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6240
         Picture         =   "frmcsms_selectestimate.frx":05AC
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2040
         Width           =   405
      End
      Begin VB.TextBox txtstdrate 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   5490
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2010
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Job Category"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   570
         TabIndex        =   28
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Job Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   3
         Left            =   375
         TabIndex        =   27
         Top             =   900
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Flat Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   4
         Left            =   975
         TabIndex        =   26
         Top             =   2070
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Standard Time Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   6
         Left            =   3840
         TabIndex        =   25
         Top             =   2070
         Width           =   1605
      End
      Begin VB.Label Label1 
         Caption         =   "Note/ Suggested Jobs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   7
         Left            =   240
         TabIndex        =   24
         Top             =   2970
         Width           =   1425
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
         TabIndex        =   23
         Top             =   2460
         Width           =   225
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   945
         TabIndex        =   22
         Top             =   2490
         Width           =   720
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Charge To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4590
         TabIndex        =   21
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label labDetCost 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Job Cost"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   930
         TabIndex        =   20
         Top             =   1650
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   855
      Left            =   6120
      MouseIcon       =   "frmcsms_selectestimate.frx":0B36
      MousePointer    =   99  'Custom
      Picture         =   "frmcsms_selectestimate.frx":0C88
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Save Entry"
      Top             =   4890
      Width           =   735
   End
   Begin VB.TextBox txtAppt 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1710
      TabIndex        =   5
      Top             =   5100
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtSaveorEdit 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   1305
   End
   Begin VB.TextBox txtCustomer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   80
      Width           =   4635
   End
   Begin VB.TextBox txtROno 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   80
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   855
      Left            =   6840
      MouseIcon       =   "frmcsms_selectestimate.frx":0F23
      MousePointer    =   99  'Custom
      Picture         =   "frmcsms_selectestimate.frx":1075
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cancel"
      Top             =   4890
      Width           =   735
   End
   Begin VB.TextBox txtCheckMe 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Customer "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   30
      TabIndex        =   32
      Top             =   195
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Flat Rate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   3750
      TabIndex        =   31
      Top             =   3030
      Width           =   885
   End
   Begin VB.Label lblGJBP 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   30
      Top             =   5160
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label labPOCODE 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   330
      TabIndex        =   29
      Top             =   5910
      Width           =   2235
   End
End
Attribute VB_Name = "frmcsms_selectestimate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AUDIT_SQL                                          As String
Dim RSUPLOAD                                           As ADODB.Recordset
Dim AddorEdit                                          As String

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
Dim VSTATUS1                                           As Integer
Dim vBP_TYPE                                           As String
Dim errormsg                                           As String

Function SetJobType(XXX As String) As String
    If XXX = "PMS" Then
        SetJobType = "PMS"
    ElseIf XXX = "60" Or XXX = "99" Then
        SetJobType = "BP"
    Else
        SetJobType = "GJ"
    End If
End Function

Function GetJobLineNo(XXX As String)
    Dim rsGetLN                                        As ADODB.Recordset
    Set rsGetLN = New ADODB.Recordset
    Set rsGetLN = gconDMIS.Execute("Select CAST([LINE_NO] AS int) AS MAX_LINE_NO ,REP_OR from CSMS_Ro_Det where [REP_OR] = '" & XXX & "' AND LIVIL = '1' order by MAX_LINE_NO desc")
    If Not rsGetLN.EOF And Not rsGetLN.BOF Then
        GetJobLineNo = Format(NumericVal(rsGetLN!MAX_LINE_NO) + 1, "00")
    Else
        GetJobLineNo = "01"
    End If
    Set rsGetLN = Nothing
End Function

Function SetROType(XXX As String) As String
    Dim rsJOBS                                         As ADODB.Recordset
    Set rsJOBS = New ADODB.Recordset
    Set rsJOBS = gconDMIS.Execute("Select * FROM CSMS_JOBMAST WHERE JCODE = '" & XXX & "'")
    If Not rsJOBS.EOF And Not rsJOBS.BOF Then
        SetROType = LTrim(RTrim(Null2String(rsJOBS!MAIN_CAT)))
    End If
    Set rsJOBS = Nothing
End Function

Sub SaveRePairOrder()
    Dim JOBREP_OR                                       As String
    Dim JOBLEVEL                                        As String
    Dim JOBLINE_NO                                      As String
    Dim JOBDETCDE                                       As String
    Dim VLastUpdateTime                                 As String
    Dim JOBDETDSC                                       As String
    Dim JOBDETUNT                                       As String
    Dim VLastUpdate                                     As String
    Dim Vusercode                                       As String
    Dim JOBDETVOL                                       As Double
    Dim JOBDETPRC                                       As Double
    Dim JOBDETAMT                                       As Double
    Dim JOBCODE                                         As String
    Dim JOBWCODE                                        As String
    Dim xApptNo                                         As String
    Dim JOBTAXRATE                                      As Double
    Dim JOBDISCRATE                                     As Double
    Dim JOBTAXVAL                                       As Double
    Dim JOBDISVAL                                       As Double
    Dim JOBPOCODE                                       As String
    Dim JOBRep_Or2                                      As String
    Dim JOBDETAIL                                       As String
    Dim JOBDET_AMT                                      As Double
    Dim JOBDIS_VAL                                      As Double
    Dim JOBDISCOUNT_2                                   As Double
    Dim xFLATRATE                                       As Double
    Dim JOBREMARKS                                      As String
    Dim JOBTECHNICIAN                                   As String
    Dim JOBDET_HRS                                      As String
    Dim TheDone                                         As String
    Dim JOBDETCOST                                      As Double
    Dim BP_TYPE                                         As String
    Dim VROTYPE                                         As String
    Dim vJobType                                        As String
    
    JOBDISVAL = 0: JOBTAXVAL = 0: JOBDETAMT = 0
    JOBDIS_VAL = 0: JOBDISCOUNT_2 = 0: JOBDISCRATE = 0

    JOBREP_OR = N2Str2Null(txtROno.Text)
    JOBLEVEL = "'1'"
    JOBLINE_NO = N2Str2Null(GetJobLineNo(txtROno))
    If labPOCODE.Caption = "" Then
        VROTYPE = N2Str2Null(SetROType(txtjCode.Text))
    Else
        VROTYPE = N2Str2Null(labPOCODE.Caption)
    End If

    If cboBP_TYPE.Visible = False Then
        BP_TYPE = N2Str2Null("")
    Else
        If cboBP_TYPE.Text = "Major" Then
            BP_TYPE = N2Str2Null("M")
        Else
            BP_TYPE = N2Str2Null("N")
        End If
    End If

    vJobType = N2Str2Null(SetJobType(SetROType(LTrim(RTrim(txtjCode.Text)))))
    JOBDETCDE = N2Str2Null(txtjCode.Text)
    JOBDETDSC = N2Str2Null(Mid(txtJobDesc.Text, 1, 500))
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
    JOBPOCODE = VROTYPE
    JOBRep_Or2 = "NULL"
    JOBDETAIL = N2Str2Null(CheckChar(txtNote.Text))
    JOBDET_AMT = JOBDETPRC
    JOBDIS_VAL = JOBDISVAL * ConvertToBIRDecimalFormat(VAT_RATE)
    JOBDISCOUNT_2 = JOBDET_AMT * JOBDISCRATE
    JOBREMARKS = N2Str2Null(CheckChar(txtNote.Text))
    JOBTECHNICIAN = "NULL"
    JOBTAXVAL = Round(((JOBDET_AMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"
    VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"
    
    If Left(txtAppt, 1) = "R" Then
        xApptNo = N2Str2Null("")
    Else
        xApptNo = N2Str2Null(txtAppt)
    End If
    
    TheDone = "'N'"
    
    SQL_STATEMENT = "insert into CSMS_RO_Det " & _
            "(TRANSTATUS, ROTYPE,JOBTYPE,ApptNo,FLATRATE,rep_or,livil,LINE_NO,detcde,detdsc,technician,det_hrs,detunt,detvol,detcost,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,USERCDE,SAVEDATE,SAVETIME,Done)" & _
            " values (" & BP_TYPE & "," & VROTYPE & "," & vJobType & "," & xApptNo & "," & xFLATRATE & "," & JOBREP_OR & ", " & JOBLEVEL & ", " & JOBLINE_NO & "," & _
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
    gconDMIS.Execute SQL_STATEMENT
    
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("AA", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(Null2String(JOBREP_OR), "REP_OR", "CSMS_REPOR"), "JOB", "RO NO: " & Null2String(JOBREP_OR), "", Null2String(vJobType))
    'NEW LOG AUDIT-----------------------------------------------------

    Dim rsRO_DET                                       As New ADODB.Recordset
    TOTJOBAMT = 0: TOTJOBDISC = 0: TOTJOBDISCVAL = 0: TOTJOBTAX = 0
    JobComTotal = 0: JobSalesTotal = 0: JobWarTotal = 0
    
    Set rsRO_DET = gconDMIS.Execute("select discount_2,det_amt,wcode,disval,taxval from CSMS_RO_Det where " & _
        " rep_or = " & JOBREP_OR & _
        " and livil = '1' " & _
        " order by LINE_NO asc")
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
    Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det where " & _
        " rep_or = " & JOBREP_OR & _
        " and livil = '2' " & _
        " order by LINE_NO asc")
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
    Set rsRO_DET = gconDMIS.Execute("select det_amt,wcode,discount_2,disval,taxval from CSMS_RO_Det " & _
        " where rep_or = " & JOBREP_OR & _
        " and livil = '3' " & _
        " order by LINE_NO asc")
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
        " where REP_OR = " & JOBREP_OR

    gconDMIS.Execute SQL_STATEMENT

    'NEW LOG AUDIT++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Dim VTRANID                                        As String
        Dim VDETID                                         As String
    
        VTRANID = FindTransactionID(N2Str2Null(txtROno), "REP_OR", "CSMS_REPOR")
        'VDETID = FindTransactionID(txtJCode, "JCODE", "CSMS_JOBMAST")
    
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, VTRANID, "R", "RO NO: " & txtROno, "", "")
    'NEW LOG AUDIT++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
End Sub

Sub CheckIFFinish()
    Dim RS                                             As New ADODB.Recordset
    Dim theRo                                          As String
    theRo = Trim(txtROno.Text)

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT Status FROM CSMS_repairOrder Where Status = 'Finish Job' AND Ro_no='" & theRo & "'")
    If Not RS.EOF And Not RS.BOF Then
        gconDMIS.Execute "UPDATE CSMS_repairOrder SET dateFinish=NULL,jstatus='W',status='Working' where Ro_no='" & theRo & "'"
        MsgBox "New Job has Been Added.Pls be advice to assigned a Technician", vbInformation, "Confirm"
    End If
    Set RS = Nothing
End Sub

Private Sub cmdCancel_Click()
    If frmcsms_selectestimate Is Nothing Then
        frmcsms_selectestimate.Show
    Else
        frmcsms_selectestimate.WindowState = 0
        frmcsms_selectestimate.ZOrder 0
    End If
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If lblGJBP.Caption = "BP" Then
        If cboBP_TYPE.Text = "" Then
            ShowIsRequiredMsg "Please Select a BP Type"
            cboBP_TYPE.SetFocus
            Exit Sub
        End If
    End If

    If NumericVal(txtstdrate) < 0.1 Then
        On Error Resume Next
        MsgBox "Std. Rate cannot be Blank/Zero or Less Than 6 Minute"
        txtstdrate.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If

    If cboBP_TYPE.Visible = True Then
        If cboBP_TYPE.Text = "Major" Then
            vBP_TYPE = "M"
        Else
            vBP_TYPE = "N"
        End If
    Else
        vBP_TYPE = ""
    End If
    Dim rsCheckRO_Det                                  As New ADODB.Recordset
    Set rsCheckRO_Det = gconDMIS.Execute("Select DETCDE from CSMS_RO_DET where LIVIL = '1' and REP_OR = '" & txtROno.Text & "' and DETCDE = '" & txtOPCODE & "'")
    If Not rsCheckRO_Det.EOF And Not rsCheckRO_Det.BOF Then
        Screen.MousePointer = 0
        MsgBox "Warning: System Does not allow Adding of Duplicate Job Codes!" & vbCrLf & "Pls select a different Job Codes to Add...", vbCritical, "Duplicates not allowed!"
        Exit Sub
    End If

    gconDMIS.BeginTrans
    If save = False Then
        MsgBox errormsg
        gconDMIS.RollbackTrans
        Exit Sub
    Else
        gconDMIS.CommitTrans
        Call ShowSuccessFullyAdded
        frmcsms_selectestimate.cmdCancel.Value = True
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdunloackFR_Click()
    If txtflatrate.Locked = True Then
        txtflatrate.Locked = False
        txtflatrate.BackColor = &HFFFFFF
        txtflatrate.ForeColor = &H0&
        
        On Error Resume Next
        txtflatrate.SetFocus
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
        
        On Error Resume Next
        txtstdrate.SetFocus
    Else
        txtstdrate.Locked = True
        txtstdrate.BackColor = &HFF8080
        txtstdrate.ForeColor = &HFFFFFF
    End If
End Sub

Private Sub Command1_Click()
    Call CheckIFFinish
End Sub

Private Sub Form_Load()
    '    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    'If lblGJBP.Caption = "BP" Then cboBP_TYPE.Visible = True
    'If Not lblGJBP.Caption = "BP" Then cboBP_TYPE.Visible = False
    Call CenterMe(frmCSMS_EstimateAddJob, Me, 1)
End Sub

Private Sub lblGJBP_Change()
    Me.Caption = "Job Selected --> " & lblGJBP.Caption
End Sub

Private Sub txtflatrate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890.", KeyAscii)
    End If
End Sub

Private Sub txtJobDesc_Change()
    txtNote = txtJobDesc.Text
End Sub
'
Private Sub txtstdrate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890.", KeyAscii)
    End If
End Sub


Function save() As Boolean
On Error GoTo ivan
    Screen.MousePointer = 11

    Dim XITEMNO                                          As Integer
    Dim JOBREP_OR                                       As String
    Dim JOBLEVEL                                        As String
    Dim JOBLINE_NO                                      As String
    Dim JOBDETCDE                                       As String
    Dim VLastUpdateTime                                 As String
    Dim JOBDETDSC                                       As String
    Dim JOBDETUNT                                       As String
    Dim VLastUpdate                                     As String
    Dim Vusercode                                       As String
    Dim JOBDETVOL                                       As Double
    Dim JOBDETPRC                                       As Double
    Dim JOBDETAMT                                       As Double
    Dim JOBCODE                                         As String
    Dim JOBWCODE                                        As String
    Dim JOBTAXRATE                                      As Double
    Dim JOBDISCRATE                                     As Double
    Dim JOBTAXVAL                                       As Double
    Dim JOBDISVAL                                       As Double
    Dim JOBPOCODE                                       As String
    Dim JOBRep_Or2                                      As String
    Dim JOBDETAIL                                       As String
    Dim JOBDET_AMT                                      As Double
    Dim JOBDIS_VAL                                      As Double
    Dim JOBDISCOUNT_2                                   As Double
    Dim xFLATRATE                                       As Double
    Dim JOBREMARKS                                      As String
    Dim JOBTECHNICIAN                                   As String
    Dim JOBDET_HRS                                      As String
    Dim xJobType                                        As String
    Dim X                                               As Long
    Dim BP_TYPE                                         As String
    Dim xrotype                                         As String
    Dim QUICK_SERVICE                                   As String
    Dim PMS_READING                                     As Long
    Dim sqlcommand                                      As String
    Dim xApptNo                                         As String
    Dim xESTIMATENO                                     As String
    Dim xTransType                                      As String
    


    JOBDISVAL = 0: JOBTAXVAL = 0: JOBDETAMT = 0
    JOBDIS_VAL = 0: JOBDISCOUNT_2 = 0: JOBDISCRATE = 0
    
    XITEMNO = (gconDMIS.Execute("select count(*) as itemno from csms_ro_det where livil ='1' and rep_or= '" & txtROno.Text & "'").Fields(0).Value)


    JOBREP_OR = N2Str2Null(txtROno.Text)
    JOBLEVEL = "'1'"
    XITEMNO = XITEMNO + 1
    JOBLINE_NO = N2Str2Null(Format(XITEMNO, "00"))
    JOBDETCDE = N2Str2Null(txtjCode.Text)
    xJobType = N2Str2Null(lblGJBP.Caption)
    JOBDETDSC = Replace(N2Str2Null(txtJobDesc), vbCrLf, " ")
    xFLATRATE = NumericVal(txtflatrate.Text)
    JOBDET_HRS = NumericVal(txtstdrate.Text)
    JOBDISCRATE = NumericVal(txtJobDiscount) / 100
    JOBWCODE = N2Str2Null(cboJobChargeTo.Text)
    If N2Str2Null(cboBP_TYPE.Text) = "Major" Then
        BP_TYPE = "M"
    ElseIf N2Str2Null(cboBP_TYPE.Text) = "Minor" Then
        BP_TYPE = "N"
    Else
        BP_TYPE = "NULL"
    End If
    PMS_READING = "0.00"
    
    JOBDETUNT = "NULL"
    JOBDETVOL = NumericVal(0)
    JOBDETPRC = NumericVal(xFLATRATE) * JOBDET_HRS
    JOBCODE = "NULL"
    JOBTAXRATE = (VAT_RATE / 100)
    JOBDETAMT = JOBDETPRC / ConvertToBIRDecimalFormat(VAT_RATE)
    JOBDISVAL = (JOBDETPRC * JOBDISCRATE) - ((JOBDETPRC * JOBDISCRATE) * JOBTAXRATE)
    JOBPOCODE = "NULL"
    JOBRep_Or2 = "NULL"
    JOBDETAIL = Replace(N2Str2Null(CheckChar(txtNote.Text)), vbCrLf, " ")
    JOBDET_AMT = JOBDETPRC
    JOBDIS_VAL = JOBDISVAL * ConvertToBIRDecimalFormat(VAT_RATE)
    JOBDISCOUNT_2 = JOBDET_AMT * JOBDISCRATE
    JOBREMARKS = N2Str2Null(CheckChar(txtNote.Text))
    JOBTECHNICIAN = "NULL"
    JOBTAXVAL = Round(((JOBDET_AMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"
    VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"
    xApptNo = N2Str2Null(txtROno.Text)
    xESTIMATENO = N2Str2Null(txtROno.Text)
    xrotype = "NULL"
    xTransType = "E"


    
    sqlcommand = "Insert Into CSMS_RO_Det " & _
    " ( TRANSTATUS, ESTIMATENO, JobType, TransType, ApptNo, FLATRATE, Rep_or, Livil, LINE_NO, Detcde, Detdsc, Technician, Det_hrs, Detunt, Detvol, Detprc, Detamt, Code, Wcode, Taxrate, Discrate, Taxval, Disval, Pocode, Rep_or2, Detail, Det_amt, Dis_val, Discount_2, USERCDE, SAVEDATE, SAVETIME, rotype, PMS_READING) " & _
    " values ( " & BP_TYPE & "," & xESTIMATENO & _
    "," & xJobType & ",'" & xTransType & _
    "'," & xApptNo & "," & xFLATRATE & _
    "," & JOBREP_OR & ", " & JOBLEVEL & _
    ", " & JOBLINE_NO & "," & JOBDETCDE & _
    "," & JOBDETDSC & "," & JOBTECHNICIAN & _
    "," & JOBDET_HRS & "," & JOBDETUNT & _
    ", " & JOBDETVOL & "," & JOBDETPRC & _
    ", " & JOBDETAMT & ", " & JOBCODE & _
    ", " & JOBWCODE & ", " & (JOBTAXRATE * 100) & _
    ", " & (JOBDISCRATE * 100) & "," & JOBTAXVAL & _
    ", " & JOBDISVAL & ", " & JOBPOCODE & _
    ", " & JOBRep_Or2 & ", " & JOBDETAIL & _
    ", " & JOBDET_AMT & ", " & JOBDIS_VAL & _
    ", " & JOBDISCOUNT_2 & ", " & Vusercode & _
    ", " & VLastUpdate & ", " & VLastUpdateTime & "," & xrotype & ", " & PMS_READING & ")"
    gconDMIS.Execute (sqlcommand)

    sqlcommand = "INSERT INTO CSMS_ESTDETAILS " & _
    " ( TRANSTATUS, ESTIMATENO, JobType, TransType, ApptNo, FLATRATE, Rep_or, Livil, LINE_NO, Detcde, Detdsc, Technician, Det_hrs, Detunt, Detvol, Detprc, Detamt, Code, Wcode, Taxrate, Discrate, Taxval, Disval, Pocode, Rep_or2, Detail, Det_amt, Dis_val, Discount_2, USERCDE, SAVEDATE, SAVETIME) " & _
    " values ( " & BP_TYPE & "," & xESTIMATENO & _
    "," & xJobType & ",'" & xTransType & _
    "'," & xApptNo & "," & xFLATRATE & _
    "," & JOBREP_OR & ", " & JOBLEVEL & _
    ", " & JOBLINE_NO & "," & JOBDETCDE & _
    "," & JOBDETDSC & "," & JOBTECHNICIAN & _
    "," & JOBDET_HRS & "," & JOBDETUNT & _
    ", " & JOBDETVOL & "," & JOBDETPRC & _
    ", " & JOBDETAMT & ", " & JOBCODE & _
    ", " & JOBWCODE & ", " & (JOBTAXRATE * 100) & _
    ", " & (JOBDISCRATE * 100) & "," & JOBTAXVAL & _
    ", " & JOBDISVAL & ", " & JOBPOCODE & _
    ", " & JOBRep_Or2 & ", " & JOBDETAIL & _
    ", " & JOBDET_AMT & ", " & JOBDIS_VAL & _
    ", " & JOBDISCOUNT_2 & ", " & Vusercode & _
    ", " & VLastUpdate & ", " & VLastUpdateTime & ")"

    gconDMIS.Execute (sqlcommand)


    save = True
    Exit Function
ivan:
    errormsg = Error
    save = False
End Function

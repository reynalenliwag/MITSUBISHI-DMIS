VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCSMSPMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PMS Add Jobs"
   ClientHeight    =   7305
   ClientLeft      =   165
   ClientTop       =   420
   ClientWidth     =   6060
   ForeColor       =   &H8000000F&
   Icon            =   "FrmPMS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNote 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   5370
      Width           =   5775
   End
   Begin VB.PictureBox Picture1 
      Height          =   1125
      Left            =   150
      ScaleHeight     =   1065
      ScaleWidth      =   5685
      TabIndex        =   12
      Top             =   1200
      Width           =   5745
      Begin VB.TextBox txtTime 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   120
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "370.00"
         Top             =   450
         Width           =   855
      End
      Begin VB.TextBox txtro 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2190
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   450
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Flat Rate Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   21
         Top             =   150
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Flat Rate Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   30
         TabIndex        =   20
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Legend: R-replace, repack,repair "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2910
         TabIndex        =   19
         Top             =   150
         Width           =   2505
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "I-inspect, clean && adjust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3570
         TabIndex        =   18
         Top             =   360
         Width           =   1785
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "L-lubricate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   3570
         TabIndex        =   17
         Top             =   570
         Width           =   1755
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "T-tighten to specified torque"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   3540
         TabIndex        =   16
         Top             =   780
         Width           =   2145
      End
   End
   Begin MSComCtl2.DTPicker dtpromise 
      Height          =   225
      Left            =   4740
      TabIndex        =   10
      Top             =   2490
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   397
      _Version        =   393216
      Format          =   51118081
      CurrentDate     =   38943
   End
   Begin VB.CommandButton cmdUnselect 
      Caption         =   "Un-Select"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1470
      TabIndex        =   9
      ToolTipText     =   "Un-Select"
      Top             =   2400
      Width           =   1275
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   8
      ToolTipText     =   "Select All Items"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   150
      TabIndex        =   1
      Top             =   60
      Width           =   5745
      Begin VB.TextBox txtCheck 
         Height          =   255
         Left            =   210
         TabIndex        =   11
         Top             =   210
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.ComboBox cboModel 
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
         ItemData        =   "FrmPMS.frx":058A
         Left            =   2010
         List            =   "FrmPMS.frx":058C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   3435
      End
      Begin VB.ComboBox cbokmReading 
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
         Left            =   2010
         TabIndex        =   3
         Text            =   "cbokmReading"
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cboMonths 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4380
         TabIndex        =   2
         Text            =   "cboMonths"
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "KM Reading  ( x 1,000 )"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   660
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "Months"
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   5
         Top             =   660
         Width           =   645
      End
   End
   Begin MSComctlLib.ListView lblTech 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   2730
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   4154
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmPMS.frx":058E
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Description"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Legend"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "code"
         Object.Width           =   0
      EndProperty
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
      Left            =   5220
      MouseIcon       =   "FrmPMS.frx":06F0
      MousePointer    =   99  'Custom
      Picture         =   "FrmPMS.frx":0842
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Cancel"
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
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
      Left            =   4500
      MouseIcon       =   "FrmPMS.frx":0B80
      MousePointer    =   99  'Custom
      Picture         =   "FrmPMS.frx":0CD2
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Select"
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Note/Suggested Jobs :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   22
      Top             =   5130
      Width           =   1905
   End
End
Attribute VB_Name = "frmCSMSPMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsloadModel                         As ADODB.Recordset
Dim tempNotes                           As String

Private Sub cbokmReading_Click()
    cboMonths.Text = ""
    txtnote.Text = Trim(cboModel) & " " & cbokmReading.Text & ",000 KM Preventive Maintenance Service Schedule"
    Call processView
End Sub


Private Sub cboModel_Click()
    cboMonths.Text = ""
    cbokmReading.Text = ""
    lblTech.Sorted = False: lblTech.ListItems.Clear
    Set rsloadModel = New ADODB.Recordset
    Set rsloadModel = gconDMIS.Execute("select Model,FlatAmt from CSMS_PMS_Hd where model ='" & Trim(cboModel) & "'")
    If Not rsloadModel.EOF And Not rsloadModel.BOF Then
        txtAmount = rsloadModel![FlatAmt]
    End If
    txtnote.Text = Trim(cboModel) & " Preventive Maintenance Service Schedule"
    'tempNotes = txtnote.Text

End Sub

Private Sub cboMonths_Click()
    cbokmReading.Text = ""

    txtnote.Text = Trim(cboModel) & " " & cboMonths & "M Preventive Maintenance Service Schedule"

    processView
End Sub

Private Sub cmdSelectAll_Click()
    Dim X                               As Long
    For X = 1 To lblTech.ListItems.Count
        If IsNumeric(Trim(lblTech.ListItems(X).SubItems(1))) = True Then
            lblTech.ListItems(X).Checked = False
        Else
            lblTech.ListItems(X).Checked = True
        End If
    Next X
End Sub

Private Sub cmdUnselect_Click()
    Dim X                               As Long
    For X = 1 To lblTech.ListItems.Count
        lblTech.ListItems(X).Checked = False
    Next X
End Sub


Sub CheckIfPMSAlreadyExistOnTheListOfPMSJob(EXIST As Boolean, JobToCompare As ListView)
    Dim X                               As Integer

    For X = 1 To JobToCompare.ListItems.Count
        If cboModel.Text = JobToCompare.ListItems(X).Text Then
            EXIST = True
            Exit Sub
        End If
    Next
End Sub

Private Sub cmdSelect_Click()
    Dim X                               As Long
    Dim sw                              As Long
    Dim xxsw                            As Long
    Dim EXIST                           As Boolean

    sw = 0

    If txtCheck.Text = "AddJobs" Then
        Call CheckIfPMSAlreadyExistOnTheListOfPMSJob(EXIST, frmCSMSNewAppointment.lblJob4Service)

        If Not EXIST Then
            If lblTech.ListItems.Count = 0 Then
                MsgBox "Theres no Job for this PMS Jobs", vbInformation, "PMS JOBS"
                cboModel.SetFocus
                Exit Sub
            End If

            With frmCSMSNewAppointment.lblJob4Service
                .Sorted = False
                .ListItems.Add , , cboModel.Text
                .ListItems(.ListItems.Count).ListSubItems.Add 1, , "PMS"
                .ListItems(.ListItems.Count).ListSubItems.Add 2, , Trim(cboModel.Text) & " Preventive Maintenance Service"
                .ListItems(.ListItems.Count).ListSubItems.Add 3, , txtAmount.Text
                .ListItems(.ListItems.Count).ListSubItems.Add 4, , txtTime.Text
                .ListItems(.ListItems.Count).ListSubItems.Add 5, , "0"
                .ListItems(.ListItems.Count).ListSubItems.Add 6, , "C"
                .ListItems(.ListItems.Count).ListSubItems.Add 7, , txtnote.Text
            End With

            For X = 1 To lblTech.ListItems.Count
                If lblTech.ListItems(X).Checked = True Then
                    If IsNumeric(Trim(lblTech.ListItems(X).SubItems(1))) = False Then
                        With frmCSMSNewAppointment.lstPMSDet
                            .Sorted = False
                            .ListItems.Add , , lblTech.ListItems(X).SubItems(2)
                            .ListItems(.ListItems.Count).ListSubItems.Add 1, , "PMS"
                            .ListItems(.ListItems.Count).ListSubItems.Add 2, , lblTech.ListItems(X)
                            .ListItems(.ListItems.Count).ListSubItems.Add 3, , cboModel.Text
                        End With
                    End If
                End If
            Next X
        Else

            MsgBox "This PMS Job Already Exist on the List of Jobs", vbInformation, "PMS Jobs"
            cboModel.SetFocus
            Exit Sub
        End If

    Else
        If txtRO.Text = "" Or txtTime.Text = "" Then
            MsgBox "Please check your entries"
            Exit Sub
        End If
        Dim JOBREP_OR, JOBLEVEL, JOBLINE_NO, JOBDETCDE, VLastUpdateTime As String
        Dim JOBDETDSC, JOBDETUNT, VLastUpdate, Vusercode As String
        Dim JOBDETVOL, JOBDETPRC, JOBDETAMT As Double
        Dim JOBCODE, JOBWCODE, xApptNo  As String
        Dim JOBTAXRATE, JOBDISCRATE     As Double
        Dim JOBTAXVAL, JOBDISVAL        As Double
        Dim JOBPOCODE, JOBRep_Or2, JOBDETAIL As String
        Dim JOBDET_AMT, JOBDIS_VAL, JOBDISCOUNT_2, xFLATRATE As Double
        Dim JOBREMARKS                  As String
        Dim JOBTECHNICIAN               As String
        Dim JOBDET_HRS                  As String
        JOBDISVAL = 0: JOBTAXVAL = 0: JOBDETAMT = 0
        JOBDIS_VAL = 0: JOBDISCOUNT_2 = 0: JOBDISCRATE = 0
        xApptNo = "NULL"
        JOBLINE_NO = "0"

        '         Call CheckIfPMSAlreadyExistOnTheListOfPMSJob(EXIST, frmCSMSServiceCounter.lstJob4Service)

        If Not EXIST Then
            If lblTech.ListItems.Count = 0 Then
                MsgBox "Theres no Job for this PMS Jobs", vbInformation, "PMS JOBS"
                cboModel.SetFocus
                Exit Sub
            End If

            gconDMIS.Execute "delete from CSMS_RO_Det where ApptNo = " & xApptNo & ""
            For X = 1 To lblTech.ListItems.Count
                If lblTech.ListItems(X).Checked = True Then
                    JOBREP_OR = N2Str2Null(txtRO)
                    JOBLEVEL = "'1'"
                    JOBLINE_NO = Format(Val(JOBLINE_NO) + 1, "00")
                    JOBDETCDE = N2Str2Null(lblTech.ListItems(X).SubItems(2))
                    JOBDETDSC = N2Str2Null(Mid(lblTech.ListItems(X), 1, 500))
                    JOBDETUNT = "NULL"
                    JOBDETVOL = NumericVal(0)

                    If sw <= 0 Then
                        JOBDET_HRS = NumericVal(txtTime)
                        xFLATRATE = NumericVal(txtAmount)
                        JOBDETPRC = NumericVal(xFLATRATE) * JOBDET_HRS
                        sw = 123
                    Else
                        JOBDET_HRS = NumericVal(0)
                        xFLATRATE = NumericVal(0)
                        JOBDETPRC = NumericVal(0)
                    End If

                    JOBCODE = "NULL"
                    JOBWCODE = "NULL"
                    JOBTAXRATE = (VAT_RATE / 100)
                    JOBDISCRATE = NumericVal(0)
                    JOBDETAMT = Round(JOBDETPRC / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                    JOBDISVAL = (JOBDETPRC * JOBDISCRATE) - ((JOBDETPRC * JOBDISCRATE) * JOBTAXRATE)
                    JOBPOCODE = "NULL"
                    JOBRep_Or2 = "NULL"
                    JOBDETAIL = "NULL"
                    JOBDET_AMT = JOBDETPRC
                    JOBDIS_VAL = JOBDISVAL * ConvertToBIRDecimalFormat(VAT_RATE)
                    JOBDISCOUNT_2 = JOBDET_AMT * JOBDISCRATE
                    JOBREMARKS = "NULL"
                    JOBTECHNICIAN = "NULL"
                    JOBTAXVAL = Round(((JOBDETAMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
                    Vusercode = "" & N2Str2Null(LOGCODE) & ""
                    VLastUpdate = "'" & LOGDATE & "'"
                    VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"

                    gconDMIS.Execute "insert into CSMS_PMS_Job_Det (PMS_MODEL,rep_or,LINE_NO,detcde,detdsc) values ('Starex'," & JOBREP_OR & ", " & JOBLINE_NO & "," & JOBDETCDE & "," & JOBDETDSC & ")"

                    If xxsw <= 0 Then
                        If IsNumeric(Trim(lblTech.ListItems(X).SubItems(1))) = False Then
                            JOBPOCODE = "'PM'"
                            JOBDETCDE = N2Str2Null(cboModel.Text)
                            JOBDETDSC = N2Str2Null(Trim(cboModel.Text) & " Preventive Maintenance Service")
                            JOBDETAIL = N2Str2Null(Trim(txtnote))

                            gconDMIS.Execute "insert into CSMS_RO_Det " & _
                                             "(JOBTYPE, ApptNo,FLATRATE,rep_or,livil,LINE_NO,detcde,detdsc,technician,det_hrs,detunt,detvol,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2,USERCDE,SAVEDATE,SAVETIME)" & _
                                           " values ('PMS' ," & xApptNo & "," & xFLATRATE & "," & JOBREP_OR & ", " & JOBLEVEL & ", " & JOBLINE_NO & "," & _
                                           " " & JOBDETCDE & "," & JOBDETDSC & "," & JOBTECHNICIAN & "," & JOBDET_HRS & "," & _
                                           " " & JOBDETUNT & ", " & JOBDETVOL & "," & _
                                           " " & JOBDETPRC & ", " & JOBDETAMT & ", " & JOBCODE & _
                                             ", " & JOBWCODE & ", " & (JOBTAXRATE * 100) & ", " & (JOBDISCRATE * 100) & _
                                             ", " & JOBTAXVAL & ", " & JOBDISVAL & ", " & JOBPOCODE & _
                                             ", " & JOBRep_Or2 & ", " & JOBDETAIL & ", " & JOBDET_AMT & _
                                             ", " & JOBDIS_VAL & ", " & JOBDISCOUNT_2 & _
                                             ", " & Vusercode & _
                                             ", " & VLastUpdate & _
                                             ", " & VLastUpdateTime & ")"
                            xxsw = 123
                        End If
                    End If
                End If
            Next X
            LogAudit "A", "PMS JOB ADDED TO RO " & txtRO & " MODEL " & cboModel
        Else
            MsgBox "This PMS Job Already Exist on the List of Jobs", vbInformation, "PMS Jobs"
            cboModel.SetFocus
            Exit Sub
        End If
    End If

    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cboModel.ListIndex = -1
    cbokmReading.Text = ""
    cboMonths.Text = ""
    cbokmReading.AddItem "1"
    cbokmReading.AddItem "5"
    cbokmReading.AddItem "10"
    cbokmReading.AddItem "15"
    cbokmReading.AddItem "20"
    cbokmReading.AddItem "25"
    cbokmReading.AddItem "30"
    cbokmReading.AddItem "35"
    cbokmReading.AddItem "40"
    cbokmReading.AddItem "45"
    cbokmReading.AddItem "50"
    cbokmReading.AddItem "55"
    cbokmReading.AddItem "60"
    cbokmReading.AddItem "65"
    cbokmReading.AddItem "70"
    cbokmReading.AddItem "75"
    cbokmReading.AddItem "80"
    cbokmReading.AddItem "85"
    cbokmReading.AddItem "90"
    cbokmReading.AddItem "95"
    cbokmReading.AddItem "100"
    cboMonths.AddItem "1"
    cboMonths.AddItem "3"
    cboMonths.AddItem "6"
    cboMonths.AddItem "9"
    cboMonths.AddItem "12"
    cboMonths.AddItem "15"
    cboMonths.AddItem "18"
    cboMonths.AddItem "21"
    cboMonths.AddItem "24"
    cboMonths.AddItem "27"
    cboMonths.AddItem "30"
    cboMonths.AddItem "33"
    cboMonths.AddItem "36"
    cboMonths.AddItem "39"
    cboMonths.AddItem "42"
    cboMonths.AddItem "45"
    cboMonths.AddItem "48"
    cboMonths.AddItem "51"
    cboMonths.AddItem "54"
    cboMonths.AddItem "57"
    cboMonths.AddItem "60"
    Dim rsloadModel                     As ADODB.Recordset
    Set rsloadModel = New ADODB.Recordset
    Set rsloadModel = gconDMIS.Execute("select Model from CSMS_PMS_Hd order by Model asc")
    If Not rsloadModel.EOF And Not rsloadModel.BOF Then
        cboModel.Clear
        Do Until rsloadModel.EOF
            cboModel.AddItem rsloadModel![Model]
            rsloadModel.MoveNext
        Loop
    End If

End Sub
Sub processView()
    Dim xfield                          As String
    Dim xtime                           As String
    If cbokmReading = "1" Then
        xfield = "KM1_1"
        xtime = "1.70"
    ElseIf cbokmReading = "5" Then
        xfield = "KM5_3"
        xtime = "1.50"
    ElseIf cbokmReading = "10" Then
        xfield = "KM10_6"
        xtime = "2.50"
    ElseIf cbokmReading = "15" Then
        xfield = "KM15_9"
        xtime = "2.20"
    ElseIf cbokmReading = "20" Then
        xfield = "KM20_12"
        xtime = "6.70"
    ElseIf cbokmReading = "25" Then
        xfield = "KM25_15"
        xtime = "2.20"
    ElseIf cbokmReading = "30" Then
        xfield = "KM30_18"
        xtime = "3.80"
    ElseIf cbokmReading = "35" Then
        xfield = "KM35_21"
        xtime = "2.20"
    ElseIf cbokmReading = "40" Then
        xfield = "KM40_24"
        xtime = "10.2"
    ElseIf cbokmReading = "45" Then
        xfield = "KM45_27"
        xtime = "2.20"
    ElseIf cbokmReading = "50" Then
        xfield = "KM50_30"
        xtime = "3.50"
    ElseIf cbokmReading = "55" Then
        xfield = "KM55_33"
        xtime = "2.20"
    ElseIf cbokmReading = "60" Then
        xfield = "KM60_36"
        xtime = "7.8"
    ElseIf cbokmReading = "65" Then
        xfield = "KM65_39"
        xtime = "2.20"
    ElseIf cbokmReading = "70" Then
        xfield = "KM70_42"
        xtime = "2.50"
    ElseIf cbokmReading = "75" Then
        xfield = "KM75_45"
        xtime = "2.20"
    ElseIf cbokmReading = "80" Then
        xfield = "KM80_48"
        xtime = "11.6"
    ElseIf cbokmReading = "85" Then
        xfield = "KM85_51"
        xtime = "2.20"
    ElseIf cbokmReading = "90" Then
        xfield = "KM90_54"
        xtime = "3.8"
    ElseIf cbokmReading = "95" Then
        xfield = "KM95_57"
        xtime = "2.20"
    ElseIf cbokmReading = "100" Then
        xfield = "KM100_60"
        xtime = "7.5"
    End If

    If cboMonths = "1" Then
        xfield = "KM1_1"
        xtime = "1.7"
    ElseIf cboMonths = "3" Then
        xfield = "KM5_3"
        xtime = "1.5"
    ElseIf cboMonths = "6" Then
        xfield = "KM10_6"
        xtime = "2.5"
    ElseIf cboMonths = "9" Then
        xfield = "KM15_9"
        xtime = "2.2"
    ElseIf cboMonths = "12" Then
        xfield = "KM20_12"
        xtime = "6.7"
    ElseIf cboMonths = "15" Then
        xfield = "KM25_15"
        xtime = "2.2"
    ElseIf cboMonths = "18" Then
        xfield = "KM30_18"
        xtime = "3.8"
    ElseIf cboMonths = "21" Then
        xfield = "KM35_21"
        xtime = "2.2"
    ElseIf cboMonths = "24" Then
        xfield = "KM40_24"
        xtime = "10.2"
    ElseIf cboMonths = "27" Then
        xfield = "KM45_27"
        xtime = "2.2"
    ElseIf cboMonths = "30" Then
        xfield = "KM50_30"
        xtime = "3.5"
    ElseIf cboMonths = "33" Then
        xfield = "KM55_33"
        xtime = "2.2"
    ElseIf cboMonths = "36" Then
        xfield = "KM60_36"
        xtime = "7.8"
    ElseIf cboMonths = "39" Then
        xfield = "KM65_39"
        xtime = "2.2"
    ElseIf cboMonths = "42" Then
        xfield = "KM70_42"
        xtime = "2.5"
    ElseIf cboMonths = "45" Then
        xfield = "KM75_45"
        xtime = "2.20"
    ElseIf cboMonths = "48" Then
        xfield = "KM80_48"
        xtime = "11.6"
    ElseIf cboMonths = "51" Then
        xfield = "KM85_51"
        xtime = "2.2"
    ElseIf cboMonths = "54" Then
        xfield = "KM90_54"
        xtime = "3.8"
    ElseIf cboMonths = "57" Then
        xfield = "KM95_57"
        xtime = "2.2"
    ElseIf cboMonths = "60" Then
        xfield = "KM100_60"
        xtime = "7.5"
    End If


    txtTime.Text = xtime
    lblTech.Sorted = False: lblTech.ListItems.Clear
    Dim rsViewPMS                       As ADODB.Recordset
    Set rsViewPMS = New ADODB.Recordset
    Set rsViewPMS = gconDMIS.Execute("Select PSM_Description, " & xfield & ",code from [CSMS_PSM_DET] where " & xfield & " is not null and model = '" & Trim(cboModel) & "' order by id")
    If Not rsViewPMS.EOF And Not rsViewPMS.BOF Then
        Listview_Loadval Me.lblTech.ListItems, rsViewPMS
    End If
    cmdSelectAll.Value = True
    Dim X                               As Long
    For X = 1 To lblTech.ListItems.Count
        If IsNumeric(Trim(lblTech.ListItems(X).SubItems(1))) = True Then
            xtime = lblTech.ListItems(X).SubItems(1)
            lblTech.ListItems(X).Checked = False

        End If
    Next X

End Sub

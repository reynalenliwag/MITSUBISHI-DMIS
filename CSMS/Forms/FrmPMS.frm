VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCSMSPMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PMS Add Jobs"
   ClientHeight    =   7380
   ClientLeft      =   165
   ClientTop       =   420
   ClientWidth     =   6855
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
   Icon            =   "FrmPMS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox optQUICK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "QUICK SERVICE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2730
      MaskColor       =   &H000000FF&
      TabIndex        =   28
      Top             =   2400
      Width           =   1665
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   810
      Left            =   150
      MouseIcon       =   "FrmPMS.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "FrmPMS.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Add/Edit/Delete PMS Jobs"
      Top             =   6450
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtNote 
      Height          =   1035
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   5370
      Width           =   6615
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   150
      ScaleHeight     =   1065
      ScaleWidth      =   6525
      TabIndex        =   12
      Top             =   1200
      Width           =   6585
      Begin VB.TextBox txtTime 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1650
         TabIndex        =   15
         Top             =   180
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1650
         TabIndex        =   14
         Text            =   "420.00"
         Top             =   510
         Width           =   855
      End
      Begin VB.TextBox txtro 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   510
         Width           =   1185
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Flat Rate Time"
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
         Left            =   450
         TabIndex        =   21
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Flat Rate Amount"
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
         Left            =   210
         TabIndex        =   20
         Top             =   570
         Width           =   1395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Legend: R-replace, repack,repair "
         Height          =   210
         Left            =   3810
         TabIndex        =   19
         Top             =   150
         Width           =   2430
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I-inspect, clean && adjust"
         Height          =   210
         Left            =   4470
         TabIndex        =   18
         Top             =   360
         Width           =   1725
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "L-lubricate"
         Height          =   210
         Index           =   0
         Left            =   4470
         TabIndex        =   17
         Top             =   570
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T-tighten to specified torque"
         Height          =   210
         Index           =   1
         Left            =   4440
         TabIndex        =   16
         Top             =   780
         Width           =   2025
      End
   End
   Begin MSComCtl2.DTPicker dtpromise 
      Height          =   285
      Left            =   5280
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   503
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
      Format          =   138280961
      CurrentDate     =   38943
   End
   Begin VB.CommandButton cmdUnselect 
      Caption         =   "Un-Select"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1350
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   8
      ToolTipText     =   "Select All Items"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   150
      TabIndex        =   1
      Top             =   30
      Width           =   6585
      Begin VB.TextBox txtCheck 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3150
         TabIndex        =   11
         Top             =   690
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.ComboBox cboModel 
         Height          =   330
         ItemData        =   "FrmPMS.frx":09EF
         Left            =   2010
         List            =   "FrmPMS.frx":09F1
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   4485
      End
      Begin VB.ComboBox cbokmReading 
         Height          =   330
         Left            =   2010
         TabIndex        =   3
         Text            =   "cbokmReading"
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cboMonths 
         Height          =   330
         Left            =   5430
         TabIndex        =   2
         Text            =   "cboMonths"
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Model"
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
         Left            =   1440
         TabIndex        =   7
         Top             =   270
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "KM Reading  ( x 1,000 )"
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
         Left            =   180
         TabIndex        =   6
         Top             =   660
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Months"
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
         Index           =   2
         Left            =   4770
         TabIndex        =   5
         Top             =   660
         Width           =   630
      End
   End
   Begin MSComctlLib.ListView lblTech 
      Height          =   2355
      Left            =   120
      TabIndex        =   0
      Top             =   2730
      Width           =   6645
      _ExtentX        =   11721
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmPMS.frx":09F3
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
      Height          =   825
      Left            =   6060
      MouseIcon       =   "FrmPMS.frx":0B55
      MousePointer    =   99  'Custom
      Picture         =   "FrmPMS.frx":0CA7
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Cancel"
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   825
      Left            =   5340
      MouseIcon       =   "FrmPMS.frx":0FE5
      MousePointer    =   99  'Custom
      Picture         =   "FrmPMS.frx":1137
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Select"
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "&Add/Edit/Delete PMS"
      Height          =   210
      Left            =   975
      TabIndex        =   27
      Top             =   6750
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Note/Suggested Jobs :"
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
Dim AUDIT_SQL                                           As String
Dim rsloadModel                                         As ADODB.Recordset
Dim tempNotes                                           As String

Function GetJobLineNo(xRO_NO As Variant)
    Dim rsJobRoDet                                     As New ADODB.Recordset
    'Set rsJobRoDet = gconDMIS.Execute("Select CAST([LINE_NO] AS int) AS MAX_LINE_NO ,REP_OR from CSMS_Ro_Det where [REP_OR] = '" & XXX & "' AND LIVIL = '1' order by MAX_LINE_NO desc")
    Set rsJobRoDet = gconDMIS.Execute("Select CAST([LINE_NO] AS int) AS MAX_LINE_NO ,REP_OR from CSMS_RO_DET where " & _
        " LIVIL = '1' " & _
        " and REP_OR = '" & xRO_NO & _
        "' Order by LINE_NO DESC")
    If Not rsJobRoDet.EOF And Not rsJobRoDet.BOF Then
        GetJobLineNo = Format(NumericVal(rsJobRoDet!MAX_LINE_NO) + 1, "00")
    Else
        GetJobLineNo = "01"
    End If
    Set rsJobRoDet = Nothing
End Function

Sub CheckIfPMSAlreadyExistOnTheListOfPMSJob(EXIST As Boolean, JobToCompare As ListView)
    Dim X                                              As Integer

    For X = 1 To JobToCompare.ListItems.Count
        If cboModel.Text = JobToCompare.ListItems(X).Text Then
            EXIST = True
            Exit Sub
        End If
    Next
End Sub

Sub CheckIFFinish()
    Dim RS                                             As New ADODB.Recordset
    Dim theRo                                          As String
    theRo = Trim(txtro.Text)
    Set RS = gconDMIS.Execute("Select Case status " & _
                              " when 'Y' then 1 " & _
                              " when 'G' then 3 " & _
                              " when 'L' then 4 " & _
                              " when 'B' then 5 " & _
                              " when 'I' then 6 " & _
                              " when 'R' then 7 " & _
                              " else 2 " & _
                              " end as status1,status,rep_or from csms_ro_det where livil = 1 and rep_or = '" & theRo & "'order by status1 asc")
   If Not (RS.EOF And RS.BOF) Then
        RS.MoveFirst
        Do While Not RS.EOF
            If N2Str2Zero(RS!status1) = "1" Then
                 gconDMIS.Execute "UPDATE CSMS_repairOrder SET jstatus='F', status='Finish Job' where Ro_no='" & theRo & "'"
            ElseIf N2Str2Zero(RS!status1) = "2" Then
                 gconDMIS.Execute "UPDATE CSMS_repairOrder SET datefinish = NULL ,jstatus = NULL, status='Park' where Ro_no='" & theRo & "'"
            ElseIf N2Str2Zero(RS!status1) = "3" Then
                 gconDMIS.Execute "UPDATE CSMS_repairOrder SET datefinish = NULL ,jstatus = 'G', status='Going Home' where Ro_no='" & theRo & "'"
            ElseIf N2Str2Zero(RS!status1) = "4" Then
                 gconDMIS.Execute "UPDATE CSMS_repairOrder SET datefinish = NULL ,jstatus = 'L', status='Lunch Break' where Ro_no='" & theRo & "'"
            ElseIf N2Str2Zero(RS!status1) = "5" Then
                 gconDMIS.Execute "UPDATE CSMS_repairOrder SET datefinish = NULL ,jstatus = 'B', status='Break Time' where Ro_no='" & theRo & "'"
            ElseIf N2Str2Zero(RS!status1) = "6" Then
                gconDMIS.Execute "UPDATE CSMS_repairOrder SET datefinish = NULL ,jstatus = 'I', status='Idle Time' where Ro_no='" & theRo & "'"
            End If
            RS.MoveNext
        Loop
   End If
End Sub

Sub processView()
    Dim XFIELD                                         As String
    Dim xtime                                          As String
    If cbokmReading = "1" Then
        XFIELD = "KM1_1"
        xtime = "1.70"
    ElseIf cbokmReading = "5" Then
        XFIELD = "KM5_3"
        xtime = "1.50"
    ElseIf cbokmReading = "10" Then
        XFIELD = "KM10_6"
        xtime = "2.50"
    ElseIf cbokmReading = "15" Then
        XFIELD = "KM15_9"
        xtime = "2.20"
    ElseIf cbokmReading = "20" Then
        XFIELD = "KM20_12"
        xtime = "6.70"
    ElseIf cbokmReading = "25" Then
        XFIELD = "KM25_15"
        xtime = "2.20"
    ElseIf cbokmReading = "30" Then
        XFIELD = "KM30_18"
        xtime = "3.80"
    ElseIf cbokmReading = "35" Then
        XFIELD = "KM35_21"
        xtime = "2.20"
    ElseIf cbokmReading = "40" Then
        XFIELD = "KM40_24"
        xtime = "10.2"
    ElseIf cbokmReading = "45" Then
        XFIELD = "KM45_27"
        xtime = "2.20"
    ElseIf cbokmReading = "50" Then
        XFIELD = "KM50_30"
        xtime = "3.50"
    ElseIf cbokmReading = "55" Then
        XFIELD = "KM55_33"
        xtime = "2.20"
    ElseIf cbokmReading = "60" Then
        XFIELD = "KM60_36"
        xtime = "7.8"
    ElseIf cbokmReading = "65" Then
        XFIELD = "KM65_39"
        xtime = "2.20"
    ElseIf cbokmReading = "70" Then
        XFIELD = "KM70_42"
        xtime = "2.50"
    ElseIf cbokmReading = "75" Then
        XFIELD = "KM75_45"
        xtime = "2.20"
    ElseIf cbokmReading = "80" Then
        XFIELD = "KM80_48"
        xtime = "11.6"
    ElseIf cbokmReading = "85" Then
        XFIELD = "KM85_51"
        xtime = "2.20"
    ElseIf cbokmReading = "90" Then
        XFIELD = "KM90_54"
        xtime = "3.8"
    ElseIf cbokmReading = "95" Then
        XFIELD = "KM95_57"
        xtime = "2.20"
    ElseIf cbokmReading = "100" Then
        XFIELD = "KM100_60"
        xtime = "7.5"
    End If

    If cboMonths = "1" Then
        XFIELD = "KM1_1"
        xtime = "1.7"
    ElseIf cboMonths = "3" Then
        XFIELD = "KM5_3"
        xtime = "1.5"
    ElseIf cboMonths = "6" Then
        XFIELD = "KM10_6"
        xtime = "2.5"
    ElseIf cboMonths = "9" Then
        XFIELD = "KM15_9"
        xtime = "2.2"
    ElseIf cboMonths = "12" Then
        XFIELD = "KM20_12"
        xtime = "6.7"
    ElseIf cboMonths = "15" Then
        XFIELD = "KM25_15"
        xtime = "2.2"
    ElseIf cboMonths = "18" Then
        XFIELD = "KM30_18"
        xtime = "3.8"
    ElseIf cboMonths = "21" Then
        XFIELD = "KM35_21"
        xtime = "2.2"
    ElseIf cboMonths = "24" Then
        XFIELD = "KM40_24"
        xtime = "10.2"
    ElseIf cboMonths = "27" Then
        XFIELD = "KM45_27"
        xtime = "2.2"
    ElseIf cboMonths = "30" Then
        XFIELD = "KM50_30"
        xtime = "3.5"
    ElseIf cboMonths = "33" Then
        XFIELD = "KM55_33"
        xtime = "2.2"
    ElseIf cboMonths = "36" Then
        XFIELD = "KM60_36"
        xtime = "7.8"
    ElseIf cboMonths = "39" Then
        XFIELD = "KM65_39"
        xtime = "2.2"
    ElseIf cboMonths = "42" Then
        XFIELD = "KM70_42"
        xtime = "2.5"
    ElseIf cboMonths = "45" Then
        XFIELD = "KM75_45"
        xtime = "2.20"
    ElseIf cboMonths = "48" Then
        XFIELD = "KM80_48"
        xtime = "11.6"
    ElseIf cboMonths = "51" Then
        XFIELD = "KM85_51"
        xtime = "2.2"
    ElseIf cboMonths = "54" Then
        XFIELD = "KM90_54"
        xtime = "3.8"
    ElseIf cboMonths = "57" Then
        XFIELD = "KM95_57"
        xtime = "2.2"
    ElseIf cboMonths = "60" Then
        XFIELD = "KM100_60"
        xtime = "7.5"
    End If


    txtTime.Text = xtime
    lblTech.Sorted = False: lblTech.ListItems.Clear
    Dim rsViewPMS                                      As New ADODB.Recordset
    Set rsViewPMS = gconDMIS.Execute("Select PSM_Description, " & XFIELD & ",code from [CSMS_PSM_DET] where " & XFIELD & " is not null and model = '" & Trim(cboModel) & "' order by id")
    If Not rsViewPMS.EOF And Not rsViewPMS.BOF Then
        Listview_Loadval Me.lblTech.ListItems, rsViewPMS
    End If
    cmdSelectAll.Value = True
    Dim X                                              As Long
    For X = 1 To lblTech.ListItems.Count
        If IsNumeric(Trim(lblTech.ListItems(X).SubItems(1))) = True Then
            xtime = lblTech.ListItems(X).SubItems(1)
            lblTech.ListItems(X).Checked = False

        End If
    Next X

End Sub

Private Sub cbokmReading_Click()
    cboMonths.Text = ""
    txtNote.Text = "Perform: " & Format(NumericVal(cbokmReading.Text) * 1000, "###,###") & " " & Trim(cboModel.Text) & " Preventive Maintenance Check-Up"
    'txtnote.Text = Trim(cboModel) & " " & cbokmReading.Text & ",000 KM Preventive Maintenance Service Schedule"
    Call processView
End Sub

Private Sub cboModel_Click()
    cboMonths.Text = ""
    cbokmReading.Text = ""
    lblTech.Sorted = False: lblTech.ListItems.Clear
    Set rsloadModel = New ADODB.Recordset
    Set rsloadModel = gconDMIS.Execute("select Model,FlatAmt from CSMS_PMS_Hd where model ='" & Trim(cboModel) & "'")
    If Not rsloadModel.EOF And Not rsloadModel.BOF Then
        'txtAmount = rsloadModel![FlatAmt]
    End If
    
    'UPDATED BY: JUN ------------------------------------------------------------------
    'DATE UPDATED: 02-10-2009
    'DESCRIPTION: GET THE FLATRATE TO THE STANDARD DATA ENTRY OF FLATRATE IN ALL MAKE
    Dim rsFLATRATE                                      As New ADODB.Recordset
    Set rsFLATRATE = gconDMIS.Execute("Select FLATRATE from ALL_MAKE where CODE = 'H'")
    If Not rsFLATRATE.EOF And Not rsFLATRATE.BOF Then
        txtAmount = NumericVal(rsFLATRATE!FLATRATE)
    End If
    Set rsFLATRATE = Nothing
    'UPDATED BY: JUN ------------------------------------------------------------------
    
    txtNote.Text = "Perform: " & Format(NumericVal(cbokmReading.Text) * 1000, "###,###") & " " & Trim(cboModel.Text) & " Preventive Maintenance Check-Up"
    'txtnote.Text = Trim(cboModel) & " Preventive Maintenance Service Schedule"
    'tempNotes = txtnote.Text
End Sub

Private Sub cboMonths_Click()
    cbokmReading.Text = ""

    txtNote.Text = Trim(cboModel) & " " & cboMonths & " Month(s) Preventive Maintenance Check-Up"
    Call processView
End Sub

Private Sub cmdSelectAll_Click()
    Dim X                                              As Long
    For X = 1 To lblTech.ListItems.Count
        If IsNumeric(Trim(lblTech.ListItems(X).SubItems(1))) = True Then
            lblTech.ListItems(X).Checked = False
        Else
            lblTech.ListItems(X).Checked = True
        End If
    Next X
End Sub

Private Sub cmdUnselect_Click()
    Dim X                                              As Long
    For X = 1 To lblTech.ListItems.Count
        lblTech.ListItems(X).Checked = False
    Next X
End Sub

Private Sub cmdAdd_Click()
    'frmCSMSAddPms.Show 1
End Sub

Private Sub cmdSelect_Click()
    Dim X                                              As Long
    Dim sw                                             As Long
    Dim xxsw                                           As Long
    Dim EXIST                                          As Boolean
    Dim JOBWCODE                                       As String
    Dim IS_WARRANTY                                    As String
    Dim QUICK_SERVICE                                  As String
    Dim C                                              As Integer
    Dim i                                              As Integer
    sw = 0

    If cbokmReading.Text = "" And cboMonths.Text = "" Then
        ShowIsRequiredMsg ("Choose a Kilometer or Month")
        cbokmReading.SetFocus
        Exit Sub
    End If
'updated by: IEBV 12172010_0340pm
'description:   Cannot save if no jobe selected
    i = 0
    For C = 1 To lblTech.ListItems.Count
        If lblTech.ListItems(C).Checked = True Then
            i = i + 1
        End If
    Next C
    
    If i <> 0 Then
    
    Else
        MsgBox "No job selected.", vbCritical + vbOKOnly
        Exit Sub
    End If
'-----------------------------------------------------------------------------
    If optQUICK.Value = 1 Then
        QUICK_SERVICE = "Y"
    Else
        QUICK_SERVICE = "N"
    End If
    If txtCheck.Text = "AddJobs" Then
        Call CheckIfPMSAlreadyExistOnTheListOfPMSJob(EXIST, frmCSMSNewAppointment.lblJob4Service)

        If Not EXIST Then
            If lblTech.ListItems.Count = 0 Then
                MsgBox "Theres no Job for this PMS Jobs", vbInformation, "PMS JOBS"
                cboModel.SetFocus
                Exit Sub
            End If

            If NumericVal(cbokmReading.Text) <= 5 Then
                JOBWCODE = "'W'"
                IS_WARRANTY = "'Y'"
            Else
                JOBWCODE = "NULL"
                IS_WARRANTY = "N"
            End If

            With frmCSMSNewAppointment.lblJob4Service
                .Sorted = False
                .ListItems.Add , , cboModel.Text
                .ListItems(.ListItems.Count).ListSubItems.Add 1, , "PMS"
                .ListItems(.ListItems.Count).ListSubItems.Add 2, , Trim(cboModel.Text) & " " & cbokmReading & "T KM Preventive Maintenance Service"
                .ListItems(.ListItems.Count).ListSubItems.Add 3, , txtAmount.Text
                .ListItems(.ListItems.Count).ListSubItems.Add 4, , txtTime.Text
                .ListItems(.ListItems.Count).ListSubItems.Add 5, , "0"
                .ListItems(.ListItems.Count).ListSubItems.Add 6, , JOBWCODE
                .ListItems(.ListItems.Count).ListSubItems.Add 7, , txtNote.Text
                .ListItems(.ListItems.Count).ListSubItems.Add 8, , IS_WARRANTY
                .ListItems(.ListItems.Count).ListSubItems.Add 9, , ""
                .ListItems(.ListItems.Count).ListSubItems.Add 10, , QUICK_SERVICE
                .ListItems(.ListItems.Count).ListSubItems.Add 11, , NumericVal(cbokmReading) * 1000
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
        If txtro.Text = "" Or txtTime.Text = "" Then
            MsgBox "Please check your entries"
            Exit Sub
        End If
        Dim JOBREP_OR                                   As String
        Dim JOBLEVEL                                    As String
        Dim JOBLINE_NO                                  As String
        Dim JOBDETCDE                                   As String
        Dim VLastUpdateTime                             As String
        Dim JOBDETDSC                                   As String
        Dim JOBDETUNT                                   As String
        Dim VLastUpdate                                 As String
        Dim Vusercode                                   As String
        Dim JOBDETVOL                                   As Double
        Dim JOBDETPRC                                   As Double
        Dim JOBDETAMT                                   As Double
        Dim JOBCODE                                     As String
        Dim xApptNo                                     As String
        Dim JOBTAXRATE                                  As Double
        Dim JOBDISCRATE                                 As Double
        Dim JOBTAXVAL                                   As Double
        Dim JOBDISVAL                                   As Double
        Dim JOBPOCODE                                   As String
        Dim JOBRep_Or2                                  As String
        Dim JOBDETAIL                                   As String
        Dim JOBDET_AMT                                  As Double
        Dim JOBDIS_VAL                                  As Double
        Dim JOBDISCOUNT_2                               As Double
        Dim xFLATRATE                                   As Double
        Dim JOBREMARKS                                 As String
        Dim JOBTECHNICIAN                              As String
        Dim JOBDET_HRS                                 As Double


        JOBDISVAL = 0: JOBTAXVAL = 0: JOBDETAMT = 0
        JOBDIS_VAL = 0: JOBDISCOUNT_2 = 0: JOBDISCRATE = 0
        xApptNo = "NULL"
        JOBLINE_NO = "0"

        Call CheckIfPMSAlreadyExistOnTheListOfPMSJob(EXIST, frmCSMS_ServiceCounter.lstJob4Service)

        If EXIST = False Then
            If lblTech.ListItems.Count = 0 Then
                MsgBox "Theres no Job for this PMS Jobs", vbInformation, "PMS JOBS"
                cboModel.SetFocus
                Exit Sub
            End If

            For X = 1 To lblTech.ListItems.Count
                If lblTech.ListItems(X).Checked = True Then
                    JOBREP_OR = N2Str2Null(txtro)
                    JOBLEVEL = "'1'"
                    JOBLINE_NO = Val(JOBLINE_NO) + 1
                    JOBDETCDE = N2Str2Null(lblTech.ListItems(X).SubItems(2))
                    JOBDETDSC = N2Str2Null(Mid(lblTech.ListItems(X), 1, 500))
                    JOBDETUNT = "NULL"
                    JOBDETVOL = NumericVal(0)

                    JOBCODE = "NULL"
                    JOBWCODE = "NULL"
                    JOBTAXRATE = (VAT_RATE / 100)
                    JOBDISCRATE = NumericVal(0)
                    JOBDETAMT = Round(JOBDETPRC / ConvertToBIRDecimalFormat(VAT_RATE), 2)
                    JOBDISVAL = (JOBDETPRC * JOBDISCRATE) - ((JOBDETPRC * JOBDISCRATE) * JOBTAXRATE)
                    JOBPOCODE = "'PM'"
                    JOBRep_Or2 = "NULL"
                    JOBDETAIL = "'Perform: " & ToDoubleNumber(NumericVal(cbokmReading.Text) * 100) & " " & Trim(cboModel.Text) & " Preventive Maintenance Check-Up'"
                    JOBDET_AMT = JOBDETPRC
                    JOBDIS_VAL = JOBDISVAL * ConvertToBIRDecimalFormat(VAT_RATE)
                    JOBDISCOUNT_2 = JOBDET_AMT * JOBDISCRATE
                    JOBREMARKS = "NULL"
                    JOBTECHNICIAN = "NULL"
                    JOBTAXVAL = Round(((JOBDET_AMT - JOBDISCOUNT_2) / ConvertToBIRDecimalFormat(VAT_RATE)) * (VAT_RATE / 100), 2)
                    Vusercode = "" & N2Str2Null(LOGCODE) & ""
                    VLastUpdate = "'" & LOGDATE & "'"
                    VLastUpdateTime = "'" & Format(Now, "HH:MM:SS AM/PM") & "'"

                    gconDMIS.Execute "insert into CSMS_PMS_Job_Det " & _
                        "(JOBTYPE, PMS_MODEL, rep_or, LINE_NO, detcde, detdsc) " & _
                        " values ('PMS' " & _
                        ", " & N2Str2Null(cboModel.Text) & _
                        ", " & JOBREP_OR & _
                        ", " & JOBLINE_NO & _
                        ", " & JOBDETCDE & _
                        ", " & JOBDETDSC & ")"
                End If
            Next X
            JOBLINE_NO = N2Str2Null(GetJobLineNo(txtro))
            JOBPOCODE = "'PM'"
            JOBDETCDE = N2Str2Null(cboModel.Text)
            JOBDETDSC = N2Str2Null(Trim(cboModel.Text) & " " & cbokmReading & "T KM Preventive Maintenance Service")
            JOBDETAIL = N2Str2Null(Trim(txtNote))

            JOBDET_HRS = NumericVal(txtTime)
            xFLATRATE = NumericVal(txtAmount)
            JOBDETPRC = NumericVal(xFLATRATE) * JOBDET_HRS
            JOBDET_AMT = NumericVal(xFLATRATE) * JOBDET_HRS

            If NumericVal(cbokmReading.Text) <= 5 Then
                JOBWCODE = "'W'"
                IS_WARRANTY = "'Y'"
            Else
                JOBWCODE = "NULL"
                IS_WARRANTY = "'N'"
            End If

            AUDIT_SQL = "Insert into CSMS_RO_Det " & _
                "(QUICK_SERVICE, JOBTYPE, ApptNo, FLATRATE, Rep_or, Livil, LINE_NO, Detcde, Detdsc, Technician, Det_hrs, Detunt, Detvol, Detprc, Detamt, Code, Wcode, Taxrate, Discrate, Taxval, Disval, Pocode, Rep_or2, Detail, Det_amt, Dis_val, Discount_2, USERCDE, SAVEDATE, SAVETIME, STATUS1, PMS_READING)" & _
                " Values (" & N2Str2Null(QUICK_SERVICE) & ", 'PMS' ," & xApptNo & "," & xFLATRATE & "," & JOBREP_OR & ", " & JOBLEVEL & ", " & JOBLINE_NO & "," & _
                " " & JOBDETCDE & "," & JOBDETDSC & "," & JOBTECHNICIAN & "," & JOBDET_HRS & "," & _
                " " & JOBDETUNT & ", " & JOBDETVOL & "," & _
                " " & JOBDETPRC & ", " & JOBDETAMT & ", " & JOBCODE & _
                ", " & JOBWCODE & ", " & (JOBTAXRATE * 100) & ", " & (JOBDISCRATE * 100) & _
                ", " & JOBTAXVAL & ", " & JOBDISVAL & ", " & JOBPOCODE & _
                ", " & JOBRep_Or2 & ", " & JOBDETAIL & ", " & JOBDET_AMT & _
                ", " & JOBDIS_VAL & ", " & JOBDISCOUNT_2 & _
                ", " & Vusercode & _
                ", " & VLastUpdate & _
                ", " & VLastUpdateTime & "," & IS_WARRANTY & ", " & NumericVal(cbokmReading) * 1000 & ")"
            gconDMIS.Execute (AUDIT_SQL)

            'NEW LOG AUDIT ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                Dim VTRANID                                As String
                Dim VDETID                                 As String
    
                VTRANID = FindTransactionID(N2Str2Null(txtro), "REP_OR", "CSMS_REPOR")
                'VDETID = FindTransactionID(JOBDETCDE, "MODEL", "CSMS_PMS_HD")
    
                Call NEW_LogAudit("AA", "BILLING SYSTEM", AUDIT_SQL, VTRANID, "R", "JOB CODE: " & Null2String(JOBDETCDE), "PMS", "")
            'NEW LOG AUDIT ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            frmCSMS_ServiceCounter.Click_ScheduleGrid (txtro)
            
            MessagePop InfoFriend, "RO Information Updated", "Job Succesfully Added", 1000
        Else
            MsgBox "This PMS Job Already Exist on the List of Jobs", vbInformation, "PMS Jobs"
            cboModel.SetFocus
            Exit Sub
        End If
    End If

    Call CheckIFFinish
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
    Call initMemvars
    
    Dim rsloadModel                                    As New ADODB.Recordset
    
    Set rsloadModel = gconDMIS.Execute("select Model from CSMS_PMS_Hd order by Model asc")
    If Not rsloadModel.EOF And Not rsloadModel.BOF Then
        cboModel.Clear
        Do Until rsloadModel.EOF
            cboModel.AddItem rsloadModel![Model]
            rsloadModel.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
End Sub

Sub initMemvars()
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
End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    Else
        KeyAscii = LimitChar("1234567890.", KeyAscii)
    End If
End Sub

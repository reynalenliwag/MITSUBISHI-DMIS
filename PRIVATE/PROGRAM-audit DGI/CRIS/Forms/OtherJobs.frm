VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSOtherJobs 
   BackColor       =   &H00CCD9DF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Other Jobs"
   ClientHeight    =   8685
   ClientLeft      =   210
   ClientTop       =   540
   ClientWidth     =   8565
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00CCD9DF&
   Icon            =   "OtherJobs.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   8565
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optSUBLET 
      BackColor       =   &H0080FF80&
      Caption         =   "SUBLET REPAIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   825
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   1650
      Width           =   1635
   End
   Begin VB.OptionButton optGENERAL 
      BackColor       =   &H0080FFFF&
      Caption         =   "GENERAL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   1440
      Width           =   2145
   End
   Begin VB.OptionButton optENGINE 
      BackColor       =   &H0080FFFF&
      Caption         =   "ENGINE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1440
      Width           =   2145
   End
   Begin VB.OptionButton optFUEL 
      BackColor       =   &H0080FFFF&
      Caption         =   "FUEL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1440
      Width           =   2145
   End
   Begin VB.OptionButton optTRANSMISSION 
      BackColor       =   &H0080FFFF&
      Caption         =   "TRANSMISSION"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1860
      Width           =   2145
   End
   Begin VB.OptionButton optCHASSIS 
      BackColor       =   &H0080FFFF&
      Caption         =   "CHASSIS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1860
      Width           =   2145
   End
   Begin VB.OptionButton optBODY 
      BackColor       =   &H0080FFFF&
      Caption         =   "BODY"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1860
      Width           =   2145
   End
   Begin VB.OptionButton optTRIM 
      BackColor       =   &H0080FFFF&
      Caption         =   "TRIM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2280
      Width           =   2145
   End
   Begin VB.OptionButton optELECTRICAL 
      BackColor       =   &H0080FFFF&
      Caption         =   "ELECTRICAL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2280
      Width           =   2145
   End
   Begin VB.OptionButton optPAINT 
      BackColor       =   &H0080FFFF&
      Caption         =   "PAINT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2280
      Width           =   2145
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CCD9DF&
      Enabled         =   0   'False
      Height          =   1275
      Left            =   90
      TabIndex        =   20
      Top             =   -30
      Width           =   8385
      Begin VB.TextBox txtVehicle 
         BackColor       =   &H00CCD9DF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   6450
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   270
         Width           =   1815
      End
      Begin VB.TextBox txtCat 
         BackColor       =   &H00CCD9DF&
         Height          =   345
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Text3"
         Top             =   -570
         Width           =   2865
      End
      Begin VB.TextBox txtGroup 
         BackColor       =   &H00CCD9DF&
         Height          =   345
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text3"
         Top             =   -570
         Width           =   2355
      End
      Begin VB.TextBox txtJobDesc 
         BackColor       =   &H00CCD9DF&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   780
         Width           =   6585
      End
      Begin VB.TextBox txtCustomer 
         BackColor       =   &H00CCD9DF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   270
         Width           =   4725
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Category"
         Height          =   285
         Index           =   0
         Left            =   4110
         TabIndex        =   28
         Top             =   -540
         Width           =   1455
      End
      Begin VB.Label Make 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Group"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   27
         Top             =   -510
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Jobs Description "
         Height          =   315
         Left            =   150
         TabIndex        =   26
         Top             =   840
         Width           =   1755
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   345
         Left            =   150
         TabIndex        =   25
         Top             =   360
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdAdd 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   90
      MouseIcon       =   "OtherJobs.frx":06C2
      MousePointer    =   99  'Custom
      Picture         =   "OtherJobs.frx":0814
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Add New Record"
      Top             =   7845
      Width           =   705
   End
   Begin VB.TextBox txtKeyword 
      Height          =   390
      Left            =   1770
      TabIndex        =   8
      Top             =   2760
      Width           =   6675
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CCD9DF&
      Caption         =   "Category"
      Height          =   105
      Left            =   60
      TabIndex        =   1
      Top             =   330
      Visible         =   0   'False
      Width           =   3225
      Begin VB.TextBox txtSearch 
         Enabled         =   0   'False
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   -840
         Width           =   3045
      End
      Begin VB.OptionButton optMake 
         BackColor       =   &H00CCD9DF&
         Caption         =   "Make"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   -1200
         Width           =   1215
      End
      Begin VB.OptionButton optModel 
         BackColor       =   &H00CCD9DF&
         Caption         =   "Model"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   -1200
         Width           =   1215
      End
      Begin MSComctlLib.ListView lstCategory 
         Height          =   4755
         Left            =   30
         TabIndex        =   3
         Top             =   300
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   8387
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
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
         MouseIcon       =   "OtherJobs.frx":0B27
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Model"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Make"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Code"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CCD9DF&
      Caption         =   "Group"
      Height          =   75
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   3225
      Begin MSComctlLib.ListView lstGroup 
         Height          =   4095
         Left            =   60
         TabIndex        =   2
         Top             =   300
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   7223
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
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
         MouseIcon       =   "OtherJobs.frx":0C89
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Job Category"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin MSComctlLib.ListView lstJObs 
      Height          =   4425
      Left            =   60
      TabIndex        =   9
      Top             =   3240
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   7805
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "OtherJobs.frx":0DEB
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   " Code"
         Object.Width           =   3353
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Job Description"
         Object.Width           =   7410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Op.Time"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Flat Rate"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Job Code"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtActNo 
      BackColor       =   &H00CCD9DF&
      Height          =   375
      Left            =   6270
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "txtActNo"
      Top             =   180
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtROno 
      BackColor       =   &H00CCD9DF&
      Height          =   375
      Left            =   7410
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "txtROno"
      Top             =   180
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAppt 
      Height          =   345
      Left            =   5970
      TabIndex        =   12
      Top             =   150
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCheckMe 
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   630
      Width           =   1035
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CCD9DF&
      Height          =   60
      Left            =   -150
      TabIndex        =   29
      Top             =   1290
      Width           =   8775
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
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
      Left            =   7710
      MouseIcon       =   "OtherJobs.frx":0F4D
      MousePointer    =   99  'Custom
      Picture         =   "OtherJobs.frx":109F
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Close Window"
      Top             =   7830
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
      Left            =   6990
      MouseIcon       =   "OtherJobs.frx":13DD
      MousePointer    =   99  'Custom
      Picture         =   "OtherJobs.frx":152F
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Select"
      Top             =   7830
      Width           =   735
   End
   Begin VB.ComboBox cboModel 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1770
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtVehicleType 
      BackColor       =   &H00CCD9DF&
      Enabled         =   0   'False
      Height          =   345
      Left            =   5490
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "Text3"
      Top             =   720
      Width           =   2955
   End
   Begin VB.Label lbljobcode 
      Caption         =   "Label8"
      Height          =   285
      Left            =   1950
      TabIndex        =   32
      Top             =   7260
      Width           =   2835
   End
   Begin VB.Label lblcode 
      Caption         =   "Label7"
      Height          =   375
      Left            =   1980
      TabIndex        =   31
      Top             =   6750
      Width           =   2685
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Keyword"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   7
      Top             =   2820
      Width           =   1725
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Model"
      Height          =   315
      Left            =   450
      TabIndex        =   19
      Top             =   780
      Width           =   1275
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Type"
      Height          =   405
      Left            =   4260
      TabIndex        =   17
      Top             =   780
      Width           =   1185
   End
End
Attribute VB_Name = "frmCSMSOtherJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUpload                            As ADODB.Recordset
Dim ctl                                 As Control

Private Sub cboModel_Click()

    Call UploadCBO
End Sub

Function SetVehicleType(XXX As String) As String
    Dim rsModel                         As ADODB.Recordset
    Dim rsVehicleType                   As ADODB.Recordset
    Set rsModel = New ADODB.Recordset
    Set rsModel = gconDMIS.Execute("Select * from FLATRATE_MODEL where DESCRIPT = '" & XXX & "'")
    If Not rsModel.EOF And Not rsModel.BOF Then
        Set rsVehicleType = New ADODB.Recordset
        Set rsVehicleType = gconDMIS.Execute("Select * from FLATRATE_VEH_TYPE where CODE = " & N2Str2Null(rsModel!veh_type))
        If Not rsVehicleType.EOF And Not rsVehicleType.BOF Then
            SetVehicleType = Null2String(rsVehicleType!Description)
        Else
            SetVehicleType = ""
        End If
    Else
        SetVehicleType = ""
    End If
End Function

Function SetVehTypeCode(XXX As String) As String
    Dim rsVehType                       As ADODB.Recordset
    Set rsVehType = New ADODB.Recordset
    Set rsVehType = gconDMIS.Execute("Select * from FLATRATE_VEH_TYPE Where Description = '" & XXX & "'")
    If Not rsVehType.EOF And Not rsVehType.BOF Then
        SetVehTypeCode = Null2String(rsVehType!code)
    End If
End Function

Private Sub cmdAdd_Click()
    Screen.MousePointer = 11
    '    frmCSMSJobs.Show 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CheckIfJobAlreadyAdded(ADDED As Boolean)
    Dim x                               As Integer
    If Not frmCSMSNewAppointment.lblJob4Service.ListItems.Count = 0 Then
        For x = 1 To frmCSMSNewAppointment.lblJob4Service.ListItems.Count
            If lstJObs.SelectedItem.Text = frmCSMSNewAppointment.lblJob4Service.ListItems(x).Text Then
                ADDED = True
                Exit Sub
            End If

            x = x + 1
        Next
        ADDED = False
    Else
        ADDED = False
    End If

    Exit Sub

ALREADY_ADDED:
End Sub

Private Sub cmdSelect_Click()
    Dim ADDED                           As Boolean

    Call CheckIfJobAlreadyAdded(ADDED)

    If Not ADDED Then
        With frmCSMSJobSelected
            For Each ctl In .ControlS
                If TypeOf ctl Is TextBox Then
                    ctl.Text = ""
                End If
            Next ctl
            .cboJobChargeTo.Clear
            .cboJobChargeTo.AddItem "W"
            .cboJobChargeTo.AddItem "S"
            .cboJobChargeTo.AddItem "C"

            .txtCustomer = txtCustomer
            .txtROno = txtROno
            .txtAppt = txtROno
            .txtJobCat = "OTHER JOBS"
            .txtJobDesc = txtJobDesc

            .txtjCode = lstJObs.SelectedItem.SubItems(4)
            .txtFlatrate = lstJObs.SelectedItem.SubItems(3)
            .txtstdrate = lstJObs.SelectedItem.SubItems(2)
            .txtOPCODE = lstJObs.SelectedItem
            .txtSaveorEdit = "ADD"
            .txtCheckMe = txtCheckMe
            If IsBodyOrSublet(Trim(lstJObs.SelectedItem.SubItems(4))) = True Then
                .txtDetCost.Visible = True
                .labDetCost.Visible = True
            Else
                .txtDetCost.Visible = False
                .labDetCost.Visible = False
            End If
            LogAudit "A", "NEW JOB ADDED TO RO " & txtROno, " JOB CODE " & lstJObs.SelectedItem.SubItems(4)
            .Show 1
            Unload Me
        End With
    Else
        MsgBox "Job Already Added", vbInformation, "Add Jobs"
        On Error Resume Next
        lstJObs.SetFocus
    End If
End Sub

Function IsBodyOrSublet(XXX As String) As Boolean
    Dim rsJOBS                          As ADODB.Recordset
    Set rsJOBS = New ADODB.Recordset
    Set rsJOBS = gconDMIS.Execute("Select * from CSMS_Jobs Where JCode = '" & XXX & "'")
    If Not rsJOBS.EOF And Not rsJOBS.BOF Then
        If Trim(Null2String(rsJOBS!MAIN_CAT)) = "60" Or Trim(Null2String(rsJOBS!MAIN_CAT)) = "99" Or Left(Trim(Null2String(rsJOBS!JCode)), 2) = "SR" Then
            IsBodyOrSublet = True
        Else
            IsBodyOrSublet = False
        End If
    End If
End Function

Private Sub Form_Load()
    lstGroup.ListItems.Clear: lstCategory.ListItems.Clear: lstJObs.ListItems.Clear
    lstGroup.Enabled = False: lstCategory.Enabled = False: lstJObs.Enabled = False
    cmdSelect.Enabled = False
    Dim rsModel                         As ADODB.Recordset
    Set rsModel = New ADODB.Recordset
    Set rsModel = gconDMIS.Execute("Select * from FLATRATE_MODEL order by DESCRIPT asc")


    If Not rsModel.EOF And Not rsModel.BOF Then
        rsModel.MoveFirst: cboModel.Clear
        Do While Not rsModel.EOF
            cboModel.AddItem Null2String(rsModel!descript)
            rsModel.MoveNext
        Loop
    End If
    optModel.Value = True
    With frmCSMSReqJobs
        For Each ctl In .ControlS
            If TypeOf ctl Is TextBox Then
                ctl.Text = ""
            End If
        Next ctl
    End With
    optGENERAL.Value = True
    Call ShowOthJobs
End Sub

Sub UploadCBO()
    Set rsUpload = New ADODB.Recordset
    Set rsUpload = gconDMIS.Execute("Select [Desc],jCat from FLATRATE_GROUPS where VEH_TYPE = '" & SetVehTypeCode(txtVehicleType.Text) & "' Order by [Jcat] Asc")
    lstGroup.Sorted = False: lstGroup.ListItems.Clear
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Listview_Loadval Me.lstGroup.ListItems, rsUpload
        lstGroup.Enabled = True
    End If
End Sub

Private Sub lstGroup_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Screen.MousePointer = 11
    Set rsUpload = New ADODB.Recordset
    Set rsUpload = gconDMIS.Execute("Select CODE, DESCRIPTION from FLATRATE_CATEGORY where GCODE = '" & lstGroup.SelectedItem.SubItems(1) & "' AND VEH_TYPE = '" & SetVehTypeCode(txtVehicleType.Text) & "' Order by CODE Asc")
    lstCategory.Sorted = False: lstCategory.ListItems.Clear
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Listview_Loadval Me.lstCategory.ListItems, rsUpload
        lstCategory.Enabled = True
    End If
    txtGroup.Text = lstGroup.SelectedItem
    Screen.MousePointer = 0
End Sub

Private Sub lstCategory_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Screen.MousePointer = 11
    lstJObs.Enabled = False: cmdSelect.Enabled = False
    Set rsUpload = New ADODB.Recordset
    Set rsUpload = gconDMIS.Execute("Select OPERATING_CODE,DESCRIPTION,OPERATING_TIME,FLATRATE,OPERATING_CODE from FLATRATE_OTHJOBS Order by OPERATING_CODE Asc")
    lstJObs.Sorted = False: lstJObs.ListItems.Clear
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Listview_Loadval Me.lstJObs.ListItems, rsUpload
        lstJObs.Enabled = True
        cmdSelect.Enabled = True
    End If
    txtCat.Text = lstCategory.SelectedItem.SubItems(1)
    Screen.MousePointer = 0
End Sub

Sub ShowOthJobs()
    Screen.MousePointer = 11
    Dim JOB_GROUP                       As String
    If optGENERAL.Value Then JOB_GROUP = "'10'"
    If optENGINE.Value Then JOB_GROUP = "'20'"
    If optFUEL.Value Then JOB_GROUP = "'30'"
    If optTRANSMISSION.Value Then JOB_GROUP = "'40'"
    If optCHASSIS.Value Then JOB_GROUP = "'50'"
    If optBODY.Value Then JOB_GROUP = "'60'"
    If optTRIM.Value Then JOB_GROUP = "'80'"
    If optELECTRICAL.Value Then JOB_GROUP = "'90'"
    If optPAINT.Value Then JOB_GROUP = "'99'"
    If optSUBLET.Value Then JOB_GROUP = "'SR'"

    lstJObs.Enabled = False: cmdSelect.Enabled = False
    Set rsUpload = New ADODB.Recordset
    Set rsUpload = gconDMIS.Execute("Select OPERATING_CODE,DESCRIPTION,OPERATING_TIME,FLATRATE,OPERATING_CODE from FLATRATE_OTHJOBS where MAIN_CAT = " & JOB_GROUP & " Order by OPERATING_CODE Asc")
    lstJObs.Sorted = False: lstJObs.ListItems.Clear
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Listview_Loadval Me.lstJObs.ListItems, rsUpload
        lstJObs.Enabled = True
        cmdSelect.Enabled = True
    End If
    Screen.MousePointer = 0
End Sub

Function SetModel(XXX As String) As String
    Dim rsModel                         As ADODB.Recordset
    Set rsModel = New ADODB.Recordset
    Set rsModel = gconDMIS.Execute("Select * from FLATRATE_MODEL where descript = '" & XXX & "'")
    If Not rsModel.EOF And Not rsModel.BOF Then
        SetModel = Null2String(rsModel!Model)
    End If
End Function

Private Sub lstJObs_DblClick()
    cmdSelect.Value = True
End Sub

Private Sub lstJObs_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtJobDesc = lstJObs.SelectedItem.SubItems(1)
End Sub

Private Sub optBODY_Click()
    ShowOthJobs
End Sub

Private Sub optCHASSIS_Click()
    ShowOthJobs
End Sub

Private Sub optELECTRICAL_Click()
    ShowOthJobs
End Sub

Private Sub optENGINE_Click()
    ShowOthJobs
End Sub

Private Sub optFUEL_Click()
    ShowOthJobs
End Sub

Private Sub optGENERAL_Click()
    ShowOthJobs
End Sub

Private Sub optPAINT_Click()
    ShowOthJobs
End Sub

Private Sub optSUBLET_Click()
    ShowOthJobs
End Sub

Private Sub optTRANSMISSION_Click()
    ShowOthJobs
End Sub

Private Sub optTRIM_Click()
    ShowOthJobs
End Sub

Private Sub txtKeyword_Change()
    Screen.MousePointer = 11
    Dim JOB_GROUP                       As String
    If optGENERAL.Value Then JOB_GROUP = "'10'"
    If optENGINE.Value Then JOB_GROUP = "'20'"
    If optFUEL.Value Then JOB_GROUP = "'30'"
    If optTRANSMISSION.Value Then JOB_GROUP = "'40'"
    If optCHASSIS.Value Then JOB_GROUP = "'50'"
    If optBODY.Value Then JOB_GROUP = "'60'"
    If optTRIM.Value Then JOB_GROUP = "'80'"
    If optELECTRICAL.Value Then JOB_GROUP = "'90'"
    If optPAINT.Value Then JOB_GROUP = "'99'"
    If optSUBLET.Value Then JOB_GROUP = "'SR'"
    Set rsUpload = New ADODB.Recordset
    Set rsUpload = gconDMIS.Execute("Select OPERATING_CODE,DESCRIPTION,OPERATING_TIME,FLATRATE,OPERATING_CODE from FLATRATE_OTHJOBS WHERE MAIN_CAT = " & JOB_GROUP & " AND DESCRIPTION LIKE '" & txtKeyword.Text & "%' Order by OPERATING_CODE Asc")
    lstJObs.Sorted = False: lstJObs.ListItems.Clear
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Listview_Loadval Me.lstJObs.ListItems, rsUpload
        lstJObs.Enabled = True
        cmdSelect.Enabled = True
    End If
    Screen.MousePointer = 0
End Sub

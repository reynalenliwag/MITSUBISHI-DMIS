VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSReqJobs 
   BackColor       =   &H00CCD9DF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Labor Time Standards"
   ClientHeight    =   9960
   ClientLeft      =   210
   ClientTop       =   540
   ClientWidth     =   11865
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
   Icon            =   "FrmReqJobs.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9960
   ScaleWidth      =   11865
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCheckMe 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   5100
      TabIndex        =   13
      Top             =   9270
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.TextBox txtVehicleType 
      BackColor       =   &H00CCD9DF&
      Enabled         =   0   'False
      Height          =   345
      Left            =   8820
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   1950
      Width           =   2955
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CCD9DF&
      Enabled         =   0   'False
      Height          =   1785
      Left            =   3420
      TabIndex        =   19
      Top             =   30
      Width           =   8385
      Begin VB.TextBox txtCat 
         BackColor       =   &H00CCD9DF&
         Height          =   345
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   780
         Width           =   2865
      End
      Begin VB.TextBox txtGroup 
         BackColor       =   &H00CCD9DF&
         Height          =   345
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   780
         Width           =   2355
      End
      Begin VB.TextBox txtJobDesc 
         BackColor       =   &H00CCD9DF&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1230
         Width           =   6585
      End
      Begin VB.TextBox txtCustomer 
         BackColor       =   &H00CCD9DF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   270
         Width           =   6585
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Category"
         Height          =   285
         Index           =   0
         Left            =   4110
         TabIndex        =   27
         Top             =   810
         Width           =   1455
      End
      Begin VB.Label Make 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Group"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   26
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Jobs Description "
         Height          =   315
         Left            =   150
         TabIndex        =   25
         Top             =   1290
         Width           =   1755
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   345
         Left            =   150
         TabIndex        =   24
         Top             =   360
         Width           =   1755
      End
   End
   Begin VB.ComboBox cboModel 
      Height          =   345
      Left            =   5100
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1950
      Width           =   2415
   End
   Begin VB.TextBox txtKeyword 
      Height          =   390
      Left            =   5100
      TabIndex        =   8
      Top             =   2550
      Width           =   6675
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CCD9DF&
      Caption         =   "Category"
      Height          =   5175
      Left            =   90
      TabIndex        =   1
      Top             =   4620
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
         Left            =   60
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
         MouseIcon       =   "FrmReqJobs.frx":06C2
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
      Height          =   4515
      Left            =   90
      TabIndex        =   0
      Top             =   30
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
         MouseIcon       =   "FrmReqJobs.frx":0824
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
      Height          =   5925
      Left            =   3390
      TabIndex        =   9
      Top             =   3030
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   10451
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
      MouseIcon       =   "FrmReqJobs.frx":0986
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   " Code"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Job Description"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Op.Time"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Flat Rate"
         Object.Width           =   2540
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00CCD9DF&
      Height          =   60
      Left            =   3180
      TabIndex        =   28
      Top             =   2400
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
      Left            =   11040
      MouseIcon       =   "FrmReqJobs.frx":0AE8
      MousePointer    =   99  'Custom
      Picture         =   "FrmReqJobs.frx":0C3A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Close Window"
      Top             =   9060
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
      Left            =   10320
      MouseIcon       =   "FrmReqJobs.frx":0F78
      MousePointer    =   99  'Custom
      Picture         =   "FrmReqJobs.frx":10CA
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Select"
      Top             =   9060
      Width           =   735
   End
   Begin VB.Label lbljobcode 
      Caption         =   "Label8"
      Height          =   285
      Left            =   5280
      TabIndex        =   31
      Top             =   8490
      Width           =   2835
   End
   Begin VB.Label lblcode 
      Caption         =   "Label7"
      Height          =   375
      Left            =   5310
      TabIndex        =   30
      Top             =   7980
      Width           =   2685
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Model"
      Height          =   315
      Left            =   3780
      TabIndex        =   18
      Top             =   2010
      Width           =   1275
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Type"
      Height          =   405
      Left            =   7590
      TabIndex        =   16
      Top             =   2010
      Width           =   1185
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
      Left            =   3480
      TabIndex        =   7
      Top             =   2610
      Width           =   1725
   End
End
Attribute VB_Name = "frmCSMSReqJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUpload                            As ADODB.Recordset
Dim ctl                                 As Control

Private Sub cboModel_Click()
    Screen.MousePointer = 11
    txtVehicleType.Text = SetVehicleType(cboModel.Text)
    lstGroup.Enabled = False: lstCategory.Enabled = False: lstJobs.Enabled = False
    lstGroup.ListItems.Clear: lstCategory.ListItems.Clear: lstJobs.ListItems.Clear
    If txtVehicleType.Text <> "" Then UploadCBO
    Screen.MousePointer = 0
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

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CheckIfJobAlreadyAdded(ADDED As Boolean, CompareMe As ListView)
    Dim X                               As Integer

    If Not CompareMe.ListItems.Count = 0 Then
        For X = 1 To CompareMe.ListItems.Count
            If lstJobs.SelectedItem.Text = CompareMe.ListItems(X).Text Then
                ADDED = True
                Exit Sub
            End If
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

    'If txtCheckMe.Text = "main" Then Call CheckIfJobAlreadyAdded(ADDED, frmCSMSServiceCounter.lstJob4Service)
    'If txtCheckMe.Text = "ro" Then Call CheckIfJobAlreadyAdded(ADDED, frmCSMSNewAppointment.lblJob4Service)

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

            .txtCustomer.Text = txtCustomer
            .txtROno.Text = txtROno
            .txtAppt.Text = txtROno
            .txtJobCat.Text = txtCat
            .txtjobdesc.Text = txtjobdesc
            .txtjCode.Text = lstJobs.SelectedItem.SubItems(4)
            .txtflatrate.Text = lstJobs.SelectedItem.SubItems(3)
            .txtstdrate.Text = lstJobs.SelectedItem.SubItems(2)

            .txtOPCODE.Text = lstJobs.SelectedItem
            .txtSaveorEdit.Text = "ADD"
            .txtCheckMe.Text = txtCheckMe
            .Show 1
            Unload Me
        End With
    Else
        MsgBox "Job Already Added", vbInformation, "Add Jobs"
        On Error Resume Next
        lstJobs.SetFocus
    End If
End Sub

Private Sub Form_Load()
    lstGroup.ListItems.Clear: lstCategory.ListItems.Clear: lstJobs.ListItems.Clear
    lstGroup.Enabled = False: lstCategory.Enabled = False: lstJobs.Enabled = False
    txtKeyword.Enabled = False
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
    lstJobs.Enabled = False: cmdSelect.Enabled = False: txtKeyword.Enabled = False

    Set rsUpload = New ADODB.Recordset
    Set rsUpload = gconDMIS.Execute("Select OPERATING_CODE,DESCRIPTION,OPERATING_TIME,FLATRATE,OPERATING_CODE from FLATRATE_JOBS where CAT_CODE = '" & _
                                    lstCategory.SelectedItem & "' AND GCODE = '" & lstGroup.SelectedItem.SubItems(1) & "' AND VEH_TYPE = '" & _
                                    SetVehTypeCode(txtVehicleType.Text) & "' and MODEL = '" & SetModel(cboModel.Text) & "' Order by OPERATING_CODE Asc")
    lstJobs.Sorted = False: lstJobs.ListItems.Clear

    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Call Listview_Loadval(Me.lstJobs.ListItems, rsUpload)
        lstJobs.Enabled = True
        If txtCustomer.Text <> "" Then cmdSelect.Enabled = True
        txtKeyword.Enabled = True
    End If

    txtCat.Text = lstCategory.SelectedItem.SubItems(1)
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
    If cmdSelect.Enabled = True Then
        cmdSelect.Value = True
    End If
End Sub

Private Sub lstJObs_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtjobdesc = lstJobs.SelectedItem.SubItems(1)
End Sub



Private Sub txtKeyword_Change()
    Screen.MousePointer = 11
    lstJobs.Enabled = False: cmdSelect.Enabled = False
    Set rsUpload = New ADODB.Recordset
    Set rsUpload = gconDMIS.Execute("Select OPERATING_CODE,DESCRIPTION,OPERATING_TIME,FLATRATE,OPERATING_CODE from FLATRATE_JOBS where CAT_CODE = '" & lstCategory.SelectedItem & "' AND GCODE = '" & lstGroup.SelectedItem.SubItems(1) & "' AND VEH_TYPE = '" & SetVehTypeCode(txtVehicleType.Text) & "' and MODEL = '" & SetModel(cboModel.Text) & "' AND DESCRIPTION LIKE '" & txtKeyword.Text & "%' Order by OPERATING_CODE Asc")
    lstJobs.Sorted = False: lstJobs.ListItems.Clear
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Listview_Loadval Me.lstJobs.ListItems, rsUpload
        lstJobs.Enabled = True
        If txtCustomer.Text <> "" Then cmdSelect.Enabled = True
    End If
    Screen.MousePointer = 0
End Sub

Private Sub txtKeyword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Screen.MousePointer = 11
        lstJobs.Enabled = False: cmdSelect.Enabled = False
        Set rsUpload = New ADODB.Recordset
        Set rsUpload = gconDMIS.Execute("Select OPERATING_CODE,DESCRIPTION,OPERATING_TIME,FLATRATE,OPERATING_CODE from FLATRATE_JOBS where CAT_CODE = '" & lstCategory.SelectedItem & "' AND GCODE = '" & lstGroup.SelectedItem.SubItems(1) & "' AND VEH_TYPE = '" & SetVehTypeCode(txtVehicleType.Text) & "' and MODEL = '" & SetModel(cboModel.Text) & "' AND DESCRIPTION LIKE '" & txtKeyword.Text & "%' Order by OPERATING_CODE Asc")
        lstJobs.Sorted = False: lstJobs.ListItems.Clear
        If Not rsUpload.EOF And Not rsUpload.BOF Then
            Listview_Loadval Me.lstJobs.ListItems, rsUpload
            lstJobs.Enabled = True
            If txtCustomer.Text <> "" Then cmdSelect.Enabled = True
        End If
        Screen.MousePointer = 0
    End If
End Sub

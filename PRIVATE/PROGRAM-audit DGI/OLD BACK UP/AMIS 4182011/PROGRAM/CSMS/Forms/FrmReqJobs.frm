VERSION 5.00
Object = "{F985F9B0-A252-46B5-A444-E023A386B6FE}#1.0#0"; "wizBox.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSReqJobs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Labor Time Standards"
   ClientHeight    =   9840
   ClientLeft      =   210
   ClientTop       =   540
   ClientWidth     =   12960
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   12960
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      Height          =   405
      Left            =   11580
      TabIndex        =   38
      Top             =   2550
      Width           =   1245
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   10920
      Top             =   9360
   End
   Begin VB.PictureBox picShowLoading 
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   3690
      ScaleHeight     =   1185
      ScaleWidth      =   6345
      TabIndex        =   34
      Top             =   4388
      Width           =   6345
      Begin wizProgBar.Prg Prg1 
         Height          =   465
         Left            =   150
         TabIndex        =   36
         Top             =   540
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   820
         Picture         =   "FrmReqJobs.frx":06C2
         ForeColor       =   0
         BorderStyle     =   2
         BarPicture      =   "FrmReqJobs.frx":06DE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin wizBox.Box Box1 
         Height          =   1185
         Left            =   0
         Top             =   0
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   2090
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "Displaying Jobs... Please Wait..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   30
         TabIndex        =   35
         Top             =   30
         Width           =   6285
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   345
      Left            =   8640
      TabIndex        =   33
      Top             =   1950
      Width           =   255
   End
   Begin VB.ComboBox cboVehicleType 
      Height          =   345
      Left            =   6180
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   1950
      Width           =   2415
   End
   Begin VB.TextBox txtCheckMe 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   9390
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Frame Frame3 
      Enabled         =   0   'False
      Height          =   1785
      Left            =   4500
      TabIndex        =   20
      Top             =   30
      Width           =   8385
      Begin VB.TextBox txtCat 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   780
         Width           =   2865
      End
      Begin VB.TextBox txtGroup 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   780
         Width           =   2355
      End
      Begin VB.TextBox txtJobDesc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1230
         Width           =   6585
      End
      Begin VB.TextBox txtCustomer 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   270
         Width           =   6585
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Category"
         Height          =   285
         Index           =   0
         Left            =   4110
         TabIndex        =   28
         Top             =   810
         Width           =   1455
      End
      Begin VB.Label Make 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Group"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   27
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Jobs Description "
         Height          =   315
         Left            =   150
         TabIndex        =   26
         Top             =   1290
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
   Begin VB.ComboBox cboModel 
      Enabled         =   0   'False
      Height          =   345
      Left            =   10350
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1950
      Width           =   2415
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
      Left            =   4500
      MouseIcon       =   "FrmReqJobs.frx":06FA
      MousePointer    =   99  'Custom
      Picture         =   "FrmReqJobs.frx":084C
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Add New Record"
      Top             =   9075
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txtKeyword 
      Height          =   390
      Left            =   6180
      TabIndex        =   8
      Top             =   2550
      Width           =   5295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Category"
      Height          =   5175
      Left            =   90
      TabIndex        =   1
      Top             =   4620
      Width           =   4305
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
         Height          =   4455
         Left            =   60
         TabIndex        =   3
         Top             =   270
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   7858
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
         MouseIcon       =   "FrmReqJobs.frx":0B5F
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Model"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Make"
            Object.Width           =   5468
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Code"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label labDoubleClick 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "*** Double Click Category to Load Jobs ***"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   90
         TabIndex        =   37
         Top             =   4800
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Group"
      Height          =   4515
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   4305
      Begin MSComctlLib.ListView lstGroup 
         Height          =   4095
         Left            =   60
         TabIndex        =   2
         Top             =   300
         Width           =   4155
         _ExtentX        =   7329
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
         MouseIcon       =   "FrmReqJobs.frx":0CC1
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
      Left            =   4470
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
      MouseIcon       =   "FrmReqJobs.frx":0E23
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
      Left            =   4260
      TabIndex        =   29
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
      Left            =   12120
      MouseIcon       =   "FrmReqJobs.frx":0F85
      MousePointer    =   99  'Custom
      Picture         =   "FrmReqJobs.frx":10D7
      Style           =   1  'Graphical
      TabIndex        =   15
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
      Left            =   11400
      MouseIcon       =   "FrmReqJobs.frx":1415
      MousePointer    =   99  'Custom
      Picture         =   "FrmReqJobs.frx":1567
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Select"
      Top             =   9060
      Width           =   735
   End
   Begin VB.Label lbljobcode 
      Caption         =   "Label8"
      Height          =   285
      Left            =   6360
      TabIndex        =   31
      Top             =   8490
      Width           =   2835
   End
   Begin VB.Label lblcode 
      Caption         =   "Label7"
      Height          =   375
      Left            =   6390
      TabIndex        =   30
      Top             =   7980
      Width           =   2685
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Model"
      Height          =   315
      Left            =   9090
      TabIndex        =   19
      Top             =   2010
      Width           =   1275
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Type"
      Height          =   405
      Left            =   4650
      TabIndex        =   17
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
      Left            =   4560
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
Dim RSUPLOAD                                           As ADODB.Recordset
Dim CTL                                                As Control
Dim VEH_CODE, VEH_MODEL                                As String
Attribute VEH_MODEL.VB_VarUserMemId = 1073938434

Function SETVEHICLETYPE(XXX As String) As String
    Dim RSMODEL                                        As ADODB.Recordset
    Dim RSVEHICLETYPE                                  As ADODB.Recordset
    Set RSMODEL = New ADODB.Recordset
    Set RSMODEL = gconDMIS.Execute("Select * from FLATRATE_MODEL where DESCRIPT = '" & XXX & "'")
    If Not RSMODEL.EOF And Not RSMODEL.BOF Then
        Set RSVEHICLETYPE = New ADODB.Recordset
        Set RSVEHICLETYPE = gconDMIS.Execute("Select * from FLATRATE_VEH_TYPE where CODE = " & N2Str2Null(RSMODEL!VEH_TYPE))
        If Not RSVEHICLETYPE.EOF And Not RSVEHICLETYPE.BOF Then
            SETVEHICLETYPE = Null2String(RSVEHICLETYPE!Description)
        Else
            SETVEHICLETYPE = ""
        End If
    Else
        SETVEHICLETYPE = ""
    End If
End Function

Function SETVEHTYPECODE(XXX As String) As String
    Dim rsVEHICLE_TYPE                                 As ADODB.Recordset
    Set rsVEHICLE_TYPE = New ADODB.Recordset
    Set rsVEHICLE_TYPE = gconDMIS.Execute("Select * from FLATRATE_VEH_TYPE WHERE DESCRIPTION = '" & XXX & "'")
    If Not rsVEHICLE_TYPE.EOF And Not rsVEHICLE_TYPE.BOF Then
        SETVEHTYPECODE = Null2String(rsVEHICLE_TYPE!Code)
    End If
    Set rsVEHICLE_TYPE = Nothing
End Function

Function SETMODEL(XXX As String) As String
    Dim RSMODEL                                        As ADODB.Recordset
    Set RSMODEL = New ADODB.Recordset
    Set RSMODEL = gconDMIS.Execute("Select * from FLATRATE_MODEL where descript = '" & XXX & "'")
    If Not RSMODEL.EOF And Not RSMODEL.BOF Then
        SETMODEL = Null2String(RSMODEL!MODEL)
    End If
End Function

Sub InitCboVehicleType()
    Dim rsVEHICLE_TYPE                                 As ADODB.Recordset
    Set rsVEHICLE_TYPE = New ADODB.Recordset
    Set rsVEHICLE_TYPE = gconDMIS.Execute("Select * from FLATRATE_VEH_TYPE Order by Code asc")
    If Not rsVEHICLE_TYPE.EOF And Not rsVEHICLE_TYPE.BOF Then
        rsVEHICLE_TYPE.MoveFirst: cboVehicleType.Clear
        Do While Not rsVEHICLE_TYPE.EOF
            cboVehicleType.AddItem Null2String(rsVEHICLE_TYPE!Description)
            rsVEHICLE_TYPE.MoveNext
        Loop
    End If
    Set rsVEHICLE_TYPE = Nothing
End Sub

Sub UPLOADCBO()
    Set RSUPLOAD = New ADODB.Recordset
    Set RSUPLOAD = gconDMIS.Execute("Select [Desc],jCat from FLATRATE_GROUPS where VEH_TYPE = '" & SETVEHTYPECODE(cboVehicleType.Text) & "' Order by [Jcat] Asc")
    lstGroup.Sorted = False: lstGroup.ListItems.Clear
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Listview_Loadval Me.lstGroup.ListItems, RSUPLOAD
        lstGroup.Enabled = True
    End If
End Sub

Private Sub cboModel_Click()
    Screen.MousePointer = 11
    'txtVehicleType.Text = SetVehicleType(cboModel.Text)
    lstGroup.Enabled = False: lstCategory.Enabled = False: lstJObs.Enabled = False
    lstGroup.ListItems.Clear: lstCategory.ListItems.Clear: lstJObs.ListItems.Clear
    VEH_CODE = SETVEHTYPECODE(cboVehicleType.Text)
    VEH_MODEL = SETMODEL(cboModel.Text)
    If cboVehicleType.Text <> "" Then UPLOADCBO
    Screen.MousePointer = 0
End Sub

'Function SetVehTypeCode(XXX As String) As String
'    Dim rsVehType                                                     As ADODB.Recordset
'    Set rsVehType = New ADODB.Recordset
'    Set rsVehType = gconDMIS.Execute("Select * from FLATRATE_VEH_TYPE Where Description = '" & XXX & "'")
'    If Not rsVehType.EOF And Not rsVehType.BOF Then
'        SetVehTypeCode = Null2String(rsVehType!code)
'    End If
'End Function

Private Sub cboVehicleType_Click()
    Dim RSMODEL                                        As ADODB.Recordset
    Set RSMODEL = New ADODB.Recordset
    Set RSMODEL = gconDMIS.Execute("Select * from FLATRATE_MODEL WHERE VEH_TYPE = '" & SETVEHTYPECODE(cboVehicleType.Text) & "' order by DESCRIPT asc")
    If Not RSMODEL.EOF And Not RSMODEL.BOF Then
        RSMODEL.MoveFirst: cboModel.Clear
        Do While Not RSMODEL.EOF
            cboModel.AddItem Null2String(RSMODEL!DESCRIPT)
            RSMODEL.MoveNext
        Loop
    End If
    cboModel.Enabled = True
    VBComBoBoxDroppedDown cboModel
End Sub

Private Sub cmdAdd_Click()
    Screen.MousePointer = 11
    frmCSMSAddJobs.Show 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CheckIfJobAlreadyAdded(ADDED As Boolean, CompareMe As ListView)
    Dim X                                              As Integer

    If Not CompareMe.ListItems.Count = 0 Then
        For X = 1 To CompareMe.ListItems.Count
            If lstJObs.SelectedItem.Text = CompareMe.ListItems(X).Text Then
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
    Dim ADDED                                          As Boolean

    If txtCheckMe.Text = "main" Then Call CheckIfJobAlreadyAdded(ADDED, frmCSMS_ServiceCounter.lstJob4Service)
    If txtCheckMe.Text = "ro" Then Call CheckIfJobAlreadyAdded(ADDED, frmCSMSNewAppointment.lblJob4Service)

    If Not ADDED Then
        With frmCSMSJobSelected
            For Each CTL In .ControlS
                If TypeOf CTL Is TextBox Then
                    CTL.Text = ""
                End If
            Next CTL
            .cboJobChargeTo.Clear
            .cboJobChargeTo.AddItem "W"
            .cboJobChargeTo.AddItem "S"
            .cboJobChargeTo.AddItem "C"

            .lblGJBP.Caption = "GJ"
            .txtCustomer.Text = txtCustomer
            .cboBP_TYPE.Visible = False
            .txtROno.Text = txtROno
            .txtAppt.Text = txtROno
            .txtJobCat.Text = txtCat
            .txtjobdesc.Text = txtjobdesc
            .txtJCode.Text = lstJObs.SelectedItem.SubItems(4)
            .txtflatrate.Text = lstJObs.SelectedItem.SubItems(3)
            .txtstdrate.Text = lstJObs.SelectedItem.SubItems(2)

            .txtOPCODE.Text = lstJObs.SelectedItem
            .txtSaveorEdit.Text = "ADD"
            .txtCheckMe.Text = txtCheckMe
            .labPOCODE.Caption = "10"
            .Show 1
            Unload Me
            'LogAudit "A", "OTHER JOB SELECTED", "RO/JOB " & txtROno & "/" & txtJobDesc
        End With
    Else
        MsgBox "Job Already Added", vbInformation, "Add Jobs"
        On Error Resume Next
        lstJObs.SetFocus
    End If
End Sub

Private Sub Command1_Click()
    cboModel.ListIndex = -1
    cboModel.Enabled = False
    cboVehicleType.Enabled = True
    VBComBoBoxDroppedDown cboVehicleType
End Sub

Private Sub Command2_Click()
    Screen.MousePointer = 11
    lstJObs.Enabled = False: cmdSelect.Enabled = False
    Set RSUPLOAD = New ADODB.Recordset
    Set RSUPLOAD = gconDMIS.Execute("Select OPERATING_CODE,DESCRIPTION,OPERATING_TIME,FLATRATE,OPERATING_CODE from FLATRATE_JOBS where CAT_CODE = '" & lstCategory.SelectedItem & "' AND GCODE = '" & lstGroup.SelectedItem.SubItems(1) & "' AND VEH_TYPE = '" & SETVEHTYPECODE(cboVehicleType.Text) & "' and MODEL = '" & SETMODEL(cboModel.Text) & "' AND DESCRIPTION LIKE '" & txtKeyword.Text & "%' Order by OPERATING_CODE Asc")
    lstJObs.Sorted = False: lstJObs.ListItems.Clear
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Listview_Loadval Me.lstJObs.ListItems, RSUPLOAD
        lstJObs.Enabled = True
        If txtCustomer.Text <> "" Then cmdSelect.Enabled = True
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    lstGroup.ListItems.Clear: lstCategory.ListItems.Clear: lstJObs.ListItems.Clear
    lstGroup.Enabled = False: lstCategory.Enabled = False: lstJObs.Enabled = False
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtKeyword.Enabled = False
    cmdSelect.Enabled = False
    InitCboVehicleType
    optModel.Value = True
    picShowLoading.Visible = False
End Sub

Private Sub lstCategory_DblClick()
    Screen.MousePointer = 11: picShowLoading.Visible = True: Me.Enabled = False
    lstJObs.Enabled = False: cmdSelect.Enabled = False: txtKeyword.Enabled = False
    Dim rsFLATRATE_JOBS                                As ADODB.Recordset
    Dim Indx                                           As Integer
    Dim FLATRATE_DESCRIPTION                           As String
    Set RSUPLOAD = New ADODB.Recordset
    'Set rsUpload = gconDMIS.Execute("Select DISTINCT OPERATING_CODE,DESCRIPTION,OPERATING_TIME,FLATRATE,OPERATING_CODE from FLATRATE_JOBS where CAT_CODE = '" & _
     lstCategory.SelectedItem & "' AND GCODE = '" & lstGroup.SelectedItem.SubItems(1) & "' AND VEH_TYPE = '" & _
     SetVehTypeCode(cboVehicleType.Text) & "' and MODEL = '" & SetModel(cboModel.Text) & "' Order by OPERATING_CODE Asc")
    Set RSUPLOAD = gconDMIS.Execute("Select OPERATING_CODE,DESCRIPTION,OPERATING_TIME,FLATRATE,OPERATING_CODE from FLATRATE_JOBS where CAT_CODE = '" & _
                                    lstCategory.SelectedItem & "' AND GCODE = '" & lstGroup.SelectedItem.SubItems(1) & "' AND VEH_TYPE = '" & _
                                    VEH_CODE & "' and MODEL = '" & VEH_MODEL & "' AND SEQNO = '10' Order by OPERATING_CODE Asc,SEQNO ASC")
    lstJObs.Sorted = False: lstJObs.ListItems.Clear
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        RSUPLOAD.MoveFirst: Indx = 0: Prg1.Max = RSUPLOAD.RecordCount
        Do While Not RSUPLOAD.EOF
            Set rsFLATRATE_JOBS = New ADODB.Recordset
            Set rsFLATRATE_JOBS = gconDMIS.Execute("Select DESCRIPTION from FLATRATE_JOBS where CAT_CODE = '" & _
                                                   lstCategory.SelectedItem & "' AND GCODE = '" & lstGroup.SelectedItem.SubItems(1) & "' AND VEH_TYPE = '" & _
                                                   VEH_CODE & "' and MODEL = '" & VEH_MODEL & "' AND OPERATING_CODE = " & N2Str2Null(RSUPLOAD!OPERATING_CODE) & " Order by OPERATING_CODE Asc,SEQNO ASC")
            If Not rsFLATRATE_JOBS.EOF And Not rsFLATRATE_JOBS.BOF Then
                rsFLATRATE_JOBS.MoveFirst: FLATRATE_DESCRIPTION = ""
                Do While Not rsFLATRATE_JOBS.EOF
                    FLATRATE_DESCRIPTION = FLATRATE_DESCRIPTION & " " & Null2String(rsFLATRATE_JOBS!Description)
                    rsFLATRATE_JOBS.MoveNext
                Loop
                Indx = Indx + 1
                lstJObs.ListItems.Add Indx, , Null2String(RSUPLOAD!OPERATING_CODE)
                lstJObs.ListItems(Indx).ListSubItems.Add , , FLATRATE_DESCRIPTION
                lstJObs.ListItems(Indx).ListSubItems.Add , , Null2String(RSUPLOAD!OPERATING_TIME)
                lstJObs.ListItems(Indx).ListSubItems.Add , , Null2String(RSUPLOAD!FLATRATE)
                lstJObs.ListItems(Indx).ListSubItems.Add , , Null2String(RSUPLOAD!OPERATING_CODE)
            End If
            Set rsFLATRATE_JOBS = Nothing
            Prg1.Value = Indx: Prg1.Text = "Loading Available Jobs...": DoEvents
            RSUPLOAD.MoveNext
        Loop
        lstJObs.Enabled = True
        If txtCustomer.Text <> "" Then cmdSelect.Enabled = True
        txtKeyword.Enabled = True
    End If
    Set RSUPLOAD = Nothing
    txtCat.Text = lstCategory.SelectedItem.SubItems(1)
    Screen.MousePointer = 0: picShowLoading.Visible = False: Me.Enabled = True
End Sub

Private Sub lstGroup_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    Screen.MousePointer = 11
    Set RSUPLOAD = New ADODB.Recordset
    Set RSUPLOAD = gconDMIS.Execute("Select CODE, DESCRIPTION from FLATRATE_CATEGORY where GCODE = '" & lstGroup.SelectedItem.SubItems(1) & "' AND VEH_TYPE = '" & SETVEHTYPECODE(cboVehicleType.Text) & "' Order by CODE Asc")
    lstCategory.Sorted = False: lstCategory.ListItems.Clear: lstJObs.ListItems.Clear
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Listview_Loadval Me.lstCategory.ListItems, RSUPLOAD
        lstCategory.Enabled = True
    End If
    txtGroup.Text = lstGroup.SelectedItem
    Screen.MousePointer = 0
End Sub

Private Sub lstJObs_DblClick()
    If cmdSelect.Enabled = True Then cmdSelect.Value = True
End Sub

Private Sub lstJObs_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    txtjobdesc = lstJObs.SelectedItem.SubItems(1)
End Sub

Private Sub Timer1_Timer()
    If labDoubleClick.Visible = True Then
        labDoubleClick.Visible = False
    Else
        labDoubleClick.Visible = True
    End If
End Sub

Private Sub txtKeyword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Screen.MousePointer = 11
        lstJObs.Enabled = False: cmdSelect.Enabled = False
        Set RSUPLOAD = New ADODB.Recordset
        Set RSUPLOAD = gconDMIS.Execute("Select OPERATING_CODE,DESCRIPTION,OPERATING_TIME,FLATRATE,OPERATING_CODE from FLATRATE_JOBS where CAT_CODE = '" & lstCategory.SelectedItem & "' AND GCODE = '" & lstGroup.SelectedItem.SubItems(1) & "' AND VEH_TYPE = '" & SETVEHTYPECODE(cboVehicleType.Text) & "' and MODEL = '" & SETMODEL(cboModel.Text) & "' AND DESCRIPTION LIKE '" & txtKeyword.Text & "%' Order by OPERATING_CODE Asc")
        lstJObs.Sorted = False: lstJObs.ListItems.Clear
        If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
            Listview_Loadval Me.lstJObs.ListItems, RSUPLOAD
            lstJObs.Enabled = True
            If txtCustomer.Text <> "" Then cmdSelect.Enabled = True
        End If
        Screen.MousePointer = 0
    End If
End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComCtl.ocx"
Begin VB.Form frmCSMSOtherJobs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Other Jobs"
   ClientHeight    =   8220
   ClientLeft      =   210
   ClientTop       =   540
   ClientWidth     =   8775
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   8775
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Enabled         =   0   'False
      Height          =   1275
      Left            =   30
      TabIndex        =   19
      Top             =   -60
      Width           =   8715
      Begin VB.TextBox txtActNo 
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
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   6420
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   300
         Width           =   2175
      End
      Begin VB.TextBox txtVehicle 
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
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   6420
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   300
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtCat 
         BackColor       =   &H00CCD9DF&
         Height          =   345
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text3"
         Top             =   -570
         Width           =   2865
      End
      Begin VB.TextBox txtGroup 
         BackColor       =   &H00CCD9DF&
         Height          =   345
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text3"
         Top             =   -570
         Width           =   2355
      End
      Begin VB.TextBox txtJobDesc 
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
         Height          =   375
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   780
         Width           =   7035
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
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   1590
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   300
         Width           =   4755
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Category"
         Height          =   285
         Index           =   0
         Left            =   4110
         TabIndex        =   27
         Top             =   -540
         Width           =   1455
      End
      Begin VB.Label Make 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Group"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   26
         Top             =   -510
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jobs Description "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   90
         TabIndex        =   25
         Top             =   870
         Width           =   1485
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   24
         Top             =   420
         Width           =   1380
      End
   End
   Begin VB.OptionButton optSUBLET 
      Caption         =   "SUBLET REPAIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   675
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1740
      Width           =   1695
   End
   Begin VB.OptionButton optPMS 
      Caption         =   "PMS JOBS"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1320
      Width           =   1695
   End
   Begin VB.OptionButton optPAINT 
      Caption         =   "PAINT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2040
      Width           =   2295
   End
   Begin VB.OptionButton optBODY 
      Caption         =   "BODY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1680
      Width           =   2295
   End
   Begin VB.OptionButton optFUEL 
      Caption         =   "FUEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1320
      Width           =   2295
   End
   Begin VB.OptionButton optELECTRICAL 
      Caption         =   "ELECTRICAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2040
      Width           =   2295
   End
   Begin VB.OptionButton optTRIM 
      Caption         =   "TRIM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2040
      Width           =   2295
   End
   Begin VB.OptionButton optCHASSIS 
      Caption         =   "CHASSIS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1680
      Width           =   2295
   End
   Begin VB.OptionButton optENGINE 
      Caption         =   "ENGINE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1320
      Width           =   2295
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
      Height          =   810
      Left            =   90
      MouseIcon       =   "OtherJobs.frx":06C2
      MousePointer    =   99  'Custom
      Picture         =   "OtherJobs.frx":0814
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Add New Record"
      Top             =   7290
      Width           =   705
   End
   Begin VB.TextBox txtKeyword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1620
      TabIndex        =   8
      Top             =   2490
      Width           =   7065
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
      Height          =   4305
      Left            =   60
      TabIndex        =   9
      Top             =   2940
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   7594
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
   Begin VB.TextBox txtROno 
      BackColor       =   &H00CCD9DF&
      Height          =   375
      Left            =   7410
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "txtROno"
      Top             =   180
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAppt 
      Height          =   345
      Left            =   5970
      TabIndex        =   11
      Top             =   150
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCheckMe 
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   630
      Width           =   1035
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
      Height          =   825
      Left            =   7950
      MouseIcon       =   "OtherJobs.frx":0F4D
      MousePointer    =   99  'Custom
      Picture         =   "OtherJobs.frx":109F
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Close Window"
      Top             =   7290
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
      Height          =   825
      Left            =   7230
      MouseIcon       =   "OtherJobs.frx":13DD
      MousePointer    =   99  'Custom
      Picture         =   "OtherJobs.frx":152F
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Select"
      Top             =   7290
      Width           =   735
   End
   Begin VB.ComboBox cboModel 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1770
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtVehicleType 
      BackColor       =   &H00CCD9DF&
      Enabled         =   0   'False
      Height          =   345
      Left            =   5490
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "Text3"
      Top             =   720
      Width           =   2955
   End
   Begin VB.OptionButton optTRANSMISSION 
      Caption         =   "TRANSMISSION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1680
      Width           =   2295
   End
   Begin VB.OptionButton optGENERAL 
      Caption         =   "GENERAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label lblcode 
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      Height          =   255
      Left            =   180
      TabIndex        =   42
      Top             =   2970
      Width           =   2685
   End
   Begin VB.Label lbljobcode 
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label8"
      Height          =   255
      Left            =   2940
      TabIndex        =   41
      Top             =   3090
      Width           =   2685
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   150
      TabIndex        =   7
      Top             =   2550
      Width           =   1410
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Model"
      Height          =   315
      Left            =   450
      TabIndex        =   18
      Top             =   780
      Width           =   1275
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Type"
      Height          =   405
      Left            =   4260
      TabIndex        =   16
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
Dim RSUPLOAD                                           As ADODB.Recordset
Dim CTL                                                As Control

Private Sub cboModel_Click()
    Call UPLOADCBO
End Sub

Private Sub CheckIfJobAlreadyAdded(ADDED As Boolean)
    Dim X                                              As Integer
    If Not frmCSMSNewAppointment.lblJob4Service.ListItems.Count = 0 Then
        For X = 1 To frmCSMSNewAppointment.lblJob4Service.ListItems.Count
            If lstJObs.SelectedItem.Text = frmCSMSNewAppointment.lblJob4Service.ListItems(X).Text Then
                ADDED = True
                Exit Sub
            End If

            X = X + 1
        Next
        ADDED = False
    Else
        ADDED = False
    End If

    Exit Sub

ALREADY_ADDED:
End Sub

Private Sub cmdAdd_Click()
    Screen.MousePointer = 11
    frmCSMSJobs.Show 1
    
    Call txtKeyword_Change
    Screen.MousePointer = 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    If lstJObs.ListItems.Count = 0 Then Exit Sub
    
    Screen.MousePointer = 11
    Dim ADDED                                          As Boolean
    Call CheckIfJobAlreadyAdded(ADDED)

    If Not ADDED Then
        With frmCSMSJobSelected
            For Each CTL In .ControlS
                If TypeOf CTL Is TextBox Then
                    CTL.Text = ""
                End If
            Next CTL
            '†††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††
            'UPDATE BY : MJP 02-10-2008 09:39 PM
                If optBODY.Value = True Or optPAINT.Value = True Or optSUBLET.Value = True Then
                    frmCSMSJobSelected.lblGJBP.Caption = "BP"
                    frmCSMSJobSelected.cboBP_TYPE.Visible = True
                ElseIf optPMS.Value = True Then
                    frmCSMSJobSelected.lblGJBP.Caption = "PMS"
                    frmCSMSJobSelected.cboBP_TYPE.Visible = False
                Else
                    frmCSMSJobSelected.lblGJBP.Caption = "GJ"
                    frmCSMSJobSelected.cboBP_TYPE.Visible = False
                End If
            'UPDATE BY : MJP 02-10-2008 09:39 PM
            '†††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††††

            .cboJobChargeTo.Clear
            .cboJobChargeTo.AddItem ""
            .cboJobChargeTo.AddItem "W"
            .cboJobChargeTo.AddItem "S"
            .cboJobChargeTo.AddItem "C"

            .txtCustomer = txtCustomer
            .txtROno = txtROno
            .txtAppt = txtROno
            .txtJobCat = "OTHER JOBS"
            .txtJobDesc = txtJobDesc

            .txtjCode = lstJObs.SelectedItem.SubItems(4)
            .txtflatrate = lstJObs.SelectedItem.SubItems(3)
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
            Screen.MousePointer = 0
            .Show 1
            Unload Me
        End With
    Else
        Screen.MousePointer = 0
        MsgBox "Job Already Added", vbInformation, "Add Jobs"
        On Error Resume Next
        lstJObs.SetFocus
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    lstGroup.ListItems.Clear
    lstCategory.ListItems.Clear
    lstJObs.ListItems.Clear
    lstGroup.Enabled = False
    lstCategory.Enabled = False
    lstJObs.Enabled = False
    cmdSelect.Enabled = False

    Dim RSMODEL                                        As New ADODB.Recordset
    Set RSMODEL = gconDMIS.Execute("Select DESCRIPT from FLATRATE_MODEL order by DESCRIPT asc")
    cboModel.Clear
    While Not RSMODEL.EOF
        cboModel.AddItem Null2String(RSMODEL!DESCRIPT)
        RSMODEL.MoveNext
        RSMODEL.MoveNext
    Wend
    optModel.Value = True
    With frmCSMSReqJobs
        For Each CTL In .ControlS
            If TypeOf CTL Is TextBox Then
                CTL.Text = ""
            End If
        Next CTL
    End With

    optGENERAL.Value = True
    optPMS.Enabled = True
    Call SHOWOTHJOBS
End Sub

Function IsBodyOrSublet(XXX As String) As Boolean
    Dim rsJOBS                                         As New ADODB.Recordset
    Set rsJOBS = gconDMIS.Execute("Select * from CSMS_Jobs Where JCode = '" & XXX & "'")
    If Not rsJOBS.EOF And Not rsJOBS.BOF Then
        If Trim(Null2String(rsJOBS!MAIN_CAT)) = "60" Or Trim(Null2String(rsJOBS!MAIN_CAT)) = "99" Or Left(Trim(Null2String(rsJOBS!JCode)), 2) = "SR" Then
            IsBodyOrSublet = True
        Else
            IsBodyOrSublet = False
        End If
    End If
End Function

Private Sub lstCategory_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Screen.MousePointer = 11
    lstJObs.Enabled = False: cmdSelect.Enabled = False
    Set RSUPLOAD = New ADODB.Recordset
    Set RSUPLOAD = gconDMIS.Execute("Select OPERATING_CODE,DESCRIPTION,OPERATING_TIME,FLATRATE,OPERATING_CODE from FLATRATE_OTHJOBS Order by OPERATING_CODE Asc")
    lstJObs.Sorted = False: lstJObs.ListItems.Clear
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Listview_Loadval Me.lstJObs.ListItems, RSUPLOAD
        lstJObs.Enabled = True
        cmdSelect.Enabled = True
    End If
    txtCat.Text = lstCategory.SelectedItem.SubItems(1)
    Screen.MousePointer = 0
End Sub

Private Sub lstGroup_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Screen.MousePointer = 11
    Set RSUPLOAD = New ADODB.Recordset
    Set RSUPLOAD = gconDMIS.Execute("Select CODE, DESCRIPTION from FLATRATE_CATEGORY where GCODE = '" & lstGroup.SelectedItem.SubItems(1) & "' AND VEH_TYPE = '" & SETVEHTYPECODE(txtVehicleType.Text) & "' Order by CODE Asc")
    lstCategory.Sorted = False: lstCategory.ListItems.Clear
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Listview_Loadval Me.lstCategory.ListItems, RSUPLOAD
        lstCategory.Enabled = True
    End If
    txtGroup.Text = lstGroup.SelectedItem
    Screen.MousePointer = 0
End Sub

Private Sub lstJObs_DblClick()
    If lstJObs.ListItems.Count = 0 Then Exit Sub
    cmdSelect.Value = True
End Sub

Private Sub lstJObs_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtJobDesc = lstJObs.SelectedItem.SubItems(1)
End Sub

Private Sub optBODY_Click()
    SHOWOTHJOBS
End Sub

Private Sub optCHASSIS_Click()
    SHOWOTHJOBS
End Sub

Private Sub optELECTRICAL_Click()
    SHOWOTHJOBS
End Sub

Private Sub optENGINE_Click()
    SHOWOTHJOBS
End Sub

Private Sub optFUEL_Click()
    SHOWOTHJOBS
End Sub

Private Sub optGENERAL_Click()
    SHOWOTHJOBS
End Sub

Private Sub optPAINT_Click()
    SHOWOTHJOBS
End Sub

Private Sub optPMS_Click()
    SHOWOTHJOBS
End Sub

Private Sub optSUBLET_Click()
    'COMMENT BY  : MJP03122010 1038AM
    'DESCRIPTION : TO DENIED TO ALL DEALER
    '    If COMPANY_CODE = "HGC" Then
        MsgBox "This Button is been disabled", vbInformation, "Info."
        Exit Sub
    '    End If
    'COMMENT BY  : MJP03122010 1038AM
 
    Call SHOWOTHJOBS
End Sub

Private Sub optTRANSMISSION_Click()
    SHOWOTHJOBS
End Sub

Private Sub optTRIM_Click()
    SHOWOTHJOBS
End Sub

Function SETMODEL(XXX As String) As String
    Dim RSMODEL                                        As New ADODB.Recordset
    Set RSMODEL = gconDMIS.Execute("SELECT MODEL FROM FLATRATE_MODEL WHERE DESCRIPT = '" & XXX & "'")
    If Not RSMODEL.EOF And Not RSMODEL.BOF Then
        SETMODEL = Null2String(RSMODEL!Model)
    End If
End Function

Function SETVEHICLETYPE(XXX As String) As String
    Dim RSMODEL                                        As New ADODB.Recordset
    Dim RSVEHICLETYPE                                  As New ADODB.Recordset
    Set RSMODEL = gconDMIS.Execute("SELECT VEH_TYPE FROM FLATRATE_MODEL WHERE DESCRIPT = '" & XXX & "'")
    If Not RSMODEL.EOF And Not RSMODEL.BOF Then
        Set RSVEHICLETYPE = gconDMIS.Execute("SELECT DESCRIPTION FROM FLATRATE_VEH_TYPE WHERE CODE = " & N2Str2Null(RSMODEL!VEH_TYPE))
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
    Dim RSVEHTYPE                                      As New ADODB.Recordset
    Set RSVEHTYPE = gconDMIS.Execute("SELECT CODE FROM FLATRATE_VEH_TYPE WHERE DESCRIPTION = '" & XXX & "'")
    If Not RSVEHTYPE.EOF And Not RSVEHTYPE.BOF Then
        SETVEHTYPECODE = Null2String(RSVEHTYPE!Code)
    End If
End Function

Sub SHOWOTHJOBS()
    Screen.MousePointer = 11
    Dim JOB_GROUP                                      As String
    If optGENERAL.Value Then JOB_GROUP = "'10'"
    If optENGINE.Value Then JOB_GROUP = "'20'"
    If optFUEL.Value Then JOB_GROUP = "'30'"
    If optTRANSMISSION.Value Then JOB_GROUP = "'40'"
    If optCHASSIS.Value Then JOB_GROUP = "'50'"
    If optBODY.Value Then JOB_GROUP = "'60'"
    If optTRIM.Value Then JOB_GROUP = "'80'"
    If optELECTRICAL.Value Then JOB_GROUP = "'90'"
    If optPAINT.Value Then JOB_GROUP = "'99'"
    If optPMS.Value Then JOB_GROUP = "'PMS'"
    If optSUBLET.Value Then JOB_GROUP = "'SR'"

    lstJObs.Enabled = False: cmdSelect.Enabled = False
    Set RSUPLOAD = New ADODB.Recordset
    Set RSUPLOAD = gconDMIS.Execute("SELECT OPERATING_CODE,DESCRIPTION,OPERATING_TIME,FLATRATE,OPERATING_CODE x FROM FLATRATE_OTHJOBS WHERE MAIN_CAT = " & JOB_GROUP & " ORDER BY OPERATING_CODE ASC")
    lstJObs.Sorted = False: lstJObs.ListItems.Clear
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Listview_Loadval Me.lstJObs.ListItems, RSUPLOAD
        lstJObs.Enabled = True
        cmdSelect.Enabled = True
    End If
    Screen.MousePointer = 0
End Sub

Private Sub txtKeyword_Change()
    Screen.MousePointer = 11
    Dim JOB_GROUP                                      As String
    If optGENERAL.Value Then JOB_GROUP = "'10'"
    If optENGINE.Value Then JOB_GROUP = "'20'"
    If optFUEL.Value Then JOB_GROUP = "'30'"
    If optTRANSMISSION.Value Then JOB_GROUP = "'40'"
    If optCHASSIS.Value Then JOB_GROUP = "'50'"
    If optBODY.Value Then JOB_GROUP = "'60'"
    If optTRIM.Value Then JOB_GROUP = "'80'"
    If optELECTRICAL.Value Then JOB_GROUP = "'90'"
    If optPAINT.Value Then JOB_GROUP = "'99'"
    If optPMS.Value Then JOB_GROUP = "'PMS'"
    If optSUBLET.Value Then JOB_GROUP = "'SR'"
    
    Set RSUPLOAD = New ADODB.Recordset
    Set RSUPLOAD = gconDMIS.Execute("Select TOP 50 OPERATING_CODE,DESCRIPTION,OPERATING_TIME,FLATRATE,OPERATING_CODE A from FLATRATE_OTHJOBS WHERE MAIN_CAT = " & JOB_GROUP & " AND DESCRIPTION LIKE '" & Replace(txtKeyword.Text, "'", "''") & "%' Order by OPERATING_CODE Asc")
    lstJObs.Sorted = False: lstJObs.ListItems.Clear
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Listview_Loadval Me.lstJObs.ListItems, RSUPLOAD
        lstJObs.Enabled = True
        cmdSelect.Enabled = True
    End If
    Screen.MousePointer = 0
End Sub

Sub UPLOADCBO()
    Set RSUPLOAD = gconDMIS.Execute("SELECT [DESC],JCAT FROM FLATRATE_GROUPS WHERE VEH_TYPE = '" & SETVEHTYPECODE(txtVehicleType.Text) & "' ORDER BY [JCAT] ASC")
    lstGroup.Sorted = False: lstGroup.ListItems.Clear
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Listview_Loadval Me.lstGroup.ListItems, RSUPLOAD
        lstGroup.Enabled = True
    End If
End Sub

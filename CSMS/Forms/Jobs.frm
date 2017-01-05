VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSJobs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jobs Data Entry"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Jobs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   9330
   Begin VB.OptionButton optSUBLET 
      BackColor       =   &H00C0FFC0&
      Caption         =   "SUBLET"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6180
      Width           =   2955
   End
   Begin VB.OptionButton optPMS 
      BackColor       =   &H00C0FFC0&
      Caption         =   "PMS JOBS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5580
      Width           =   2955
   End
   Begin VB.OptionButton optPAINT 
      Caption         =   "PAINT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4980
      Width           =   2955
   End
   Begin VB.OptionButton optELECTRICAL 
      Caption         =   "ELECTRICAL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4380
      Width           =   2955
   End
   Begin VB.OptionButton optTRIM 
      Caption         =   "TRIM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3780
      Width           =   2955
   End
   Begin VB.OptionButton optBODY 
      Caption         =   "BODY"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3180
      Width           =   2955
   End
   Begin VB.OptionButton optCHASSIS 
      Caption         =   "CHASSIS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2580
      Width           =   2955
   End
   Begin VB.OptionButton optTRANSMISSION 
      Caption         =   "TRANSMISSION"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1950
      Width           =   2955
   End
   Begin VB.OptionButton optFUEL 
      Caption         =   "FUEL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1350
      Width           =   2955
   End
   Begin VB.OptionButton optENGINE 
      Caption         =   "ENGINE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   750
      Width           =   2955
   End
   Begin VB.OptionButton optGENERAL 
      Caption         =   "GENERAL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   150
      Width           =   2955
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   3030
      ScaleHeight     =   945
      ScaleWidth      =   6195
      TabIndex        =   13
      Top             =   6090
      Width           =   6195
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   5415
         MouseIcon       =   "Jobs.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "Jobs.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Exit Window"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4665
         MouseIcon       =   "Jobs.frx":153A
         MousePointer    =   99  'Custom
         Picture         =   "Jobs.frx":168C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Print this Record"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3915
         MouseIcon       =   "Jobs.frx":19F2
         MousePointer    =   99  'Custom
         Picture         =   "Jobs.frx":1B44
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Delete Selected Record"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3165
         MouseIcon       =   "Jobs.frx":1E6F
         MousePointer    =   99  'Custom
         Picture         =   "Jobs.frx":1FC1
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Edit Selected Record"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2415
         MouseIcon       =   "Jobs.frx":231D
         MousePointer    =   99  'Custom
         Picture         =   "Jobs.frx":246F
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Add Record"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1665
         MouseIcon       =   "Jobs.frx":2782
         MousePointer    =   99  'Custom
         Picture         =   "Jobs.frx":28D4
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Find a Record"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   915
         MouseIcon       =   "Jobs.frx":2BCE
         MousePointer    =   99  'Custom
         Picture         =   "Jobs.frx":2D20
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Move to Next Record"
         Top             =   45
         Width           =   765
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   165
         MouseIcon       =   "Jobs.frx":3078
         MousePointer    =   99  'Custom
         Picture         =   "Jobs.frx":31CA
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Move to Previous Record"
         Top             =   45
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   3150
      TabIndex        =   4
      Top             =   60
      Width           =   6105
      Begin Crystal.CrystalReport rptjobs 
         Left            =   5340
         Top             =   1110
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Jobs Master List"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.TextBox txtJCode 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1380
         TabIndex        =   0
         Top             =   240
         Width           =   2115
      End
      Begin VB.TextBox txtFlatrate 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1380
         TabIndex        =   3
         Top             =   1410
         Width           =   1065
      End
      Begin VB.TextBox txtStd_mHrs 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1380
         TabIndex        =   2
         Top             =   1020
         Width           =   1065
      End
      Begin VB.TextBox txtDesc1 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1380
         TabIndex        =   1
         Top             =   630
         Width           =   4635
      End
      Begin VB.Label labID 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Job Code"
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
         Height          =   225
         Left            =   1410
         TabIndex        =   12
         Top             =   270
         Width           =   705
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Job Code"
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
         Height          =   225
         Left            =   510
         TabIndex        =   8
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Flat Rate"
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
         Height          =   255
         Left            =   570
         TabIndex        =   7
         Top             =   1470
         Width           =   1605
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Hrs."
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
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Height          =   255
         Left            =   330
         TabIndex        =   5
         Top             =   660
         Width           =   975
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   4125
      Left            =   3150
      TabIndex        =   9
      Top             =   1920
      Width           =   6105
      Begin VB.TextBox textSearch 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   90
         MaxLength       =   35
         TabIndex        =   10
         Top             =   150
         Width           =   5925
      End
      Begin MSComctlLib.ListView lstJobs 
         Height          =   3465
         Left            =   60
         TabIndex        =   11
         Top             =   540
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   6112
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
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Jobs.frx":3529
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Job Description"
            Object.Width           =   9701
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7710
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   22
      Top             =   6180
      Width           =   1800
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
         Height          =   795
         Left            =   750
         MouseIcon       =   "Jobs.frx":368B
         MousePointer    =   99  'Custom
         Picture         =   "Jobs.frx":37DD
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   765
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   0
         MouseIcon       =   "Jobs.frx":3B1B
         MousePointer    =   99  'Custom
         Picture         =   "Jobs.frx":3C6D
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmCSMSJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsROJOBS                                           As ADODB.Recordset
Dim AddorEdit                                          As String
Dim JOB_GROUP                                          As String

Sub initMemvars()
    txtjCode.Text = ""
    txtDesc1.Text = ""
    txtStd_mHrs.Text = ""
    txtflatrate.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
        labid.Caption = rsROJOBS!ID
        txtjCode.Text = Null2String(rsROJOBS!JCode)
        txtDesc1.Text = Null2String(rsROJOBS!Description)
        txtStd_mHrs.Text = N2Str2Zero(rsROJOBS!std_mhrs)
        txtflatrate.Text = N2Str2Zero(rsROJOBS!FLATRATE)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
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
    If optPMS.Value Then JOB_GROUP = "'PMS'"
    Set rsROJOBS = New ADODB.Recordset
    rsROJOBS.Open "select * from CSMS_JobMast WHERE MAIN_CAT = " & JOB_GROUP & " order by JCode asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub RefreshRecords()
    rsRefresh
    Frame1.Enabled = False
    textSearch.Text = "":
    cmdCancel_Click
    FillGrid
End Sub

Sub FillGrid()
    Dim rsJOBS                                         As ADODB.Recordset
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
    lstJObs.Enabled = False
    lstJObs.Sorted = False: lstJObs.ListItems.Clear
    Set rsJOBS = New ADODB.Recordset
    Set rsJOBS = gconDMIS.Execute("select description,JCode from CSMS_JobMast where MAIN_CAT = " & JOB_GROUP & " order by description asc")
    If Not (rsJOBS.EOF And rsJOBS.BOF) Then
        Listview_Loadval Me.lstJObs.ListItems, rsJOBS
        lstJObs.Refresh
    End If
    lstJObs.Enabled = True
End Sub

Function GenerateJobCode(XTYPE As String) As String
    Dim rstmp As New ADODB.Recordset
    Dim JOB_DESC As String
    If JOB_GROUP = "'10'" Then JOB_DESC = "GJ"
    If JOB_GROUP = "'20'" Then JOB_DESC = "EN"
    If JOB_GROUP = "'30'" Then JOB_DESC = "FU"
    If JOB_GROUP = "'40'" Then JOB_DESC = "TR"
    If JOB_GROUP = "'50'" Then JOB_DESC = "CH"
    If JOB_GROUP = "'60'" Then JOB_DESC = "BO"
    If JOB_GROUP = "'80'" Then JOB_DESC = "TM"
    If JOB_GROUP = "'90'" Then JOB_DESC = "EL"
    If JOB_GROUP = "'99'" Then JOB_DESC = "PA"
    If JOB_GROUP = "'SR'" Then JOB_DESC = "SU"
    
    Set rstmp = gconDMIS.Execute("SELECT TOP 1 right(JCODE,4) AS NEW_CODE FROM CSMS_JOBMAST WHERE MAIN_CAT = " & JOB_GROUP & " AND LEFT(JCODE,2) = " & N2Str2Null(JOB_DESC) & " ORDER BY RIGHT(JCODE,4) DESC")
    If Not (rstmp.BOF And rstmp.EOF) Then
        GenerateJobCode = JOB_DESC & Format(TONUMERIC(rstmp!NEW_CODE) + 1, "0000")
    Else
        GenerateJobCode = JOB_DESC & Format(1, "0000")
    End If
    Set rstmp = Nothing
End Function

Sub FillSearchGrid(XXX As String)
    Dim rsJOBS                                         As ADODB.Recordset
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
    lstJObs.Enabled = False
    lstJObs.Sorted = False: lstJObs.ListItems.Clear
    Set rsJOBS = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsJOBS = gconDMIS.Execute("select description, JCode from CSMS_JobMast where MAIN_CAT = " & JOB_GROUP & " AND description like'" & XXX & "%'")
    If Not (rsJOBS.EOF And rsJOBS.BOF) Then
        Listview_Loadval Me.lstJObs.ListItems, rsJOBS
        lstJObs.Enabled = True
        lstJObs.Refresh
    End If
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "JOBS") = False Then Exit Sub
    Screen.MousePointer = 11
    rptjobs.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptjobs.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"


    PrintSQLReport rptjobs, CSMS_REPORT_PATH & "rojobs.rpt", "", CSMS_REPORT_CONNECTION, 1
    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("V", "JOBS", "", labid, "", "JOB CODE: " & txtjCode, "", "")
    'NEW LOG AUDIT-----------------------------------------------------
    'LogAudit "V", "JOBS"
    Screen.MousePointer = 0
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "JOBS") = False Then Exit Sub
    
    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    Call initMemvars
    txtjCode.Text = GenerateJobCode(JOB_GROUP)
    On Error Resume Next
    txtjCode.SetFocus
End Sub



Private Sub cmdCancel_Click()
    fraDetails.Enabled = True
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    StoreMemVars
    FillGrid
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "JOBS") = False Then Exit Sub
    On Error GoTo ErrorCode
    If Not rsROJOBS.BOF Or Not rsROJOBS.EOF Then
        If MsgBox("Are you sure?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            SQL_STATEMENT = "delete from CSMS_JobMast where id = " & labid.Caption
            gconDMIS.Execute SQL_STATEMENT

            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("X", "JOBS", SQL_STATEMENT, labid, "", "JOB CODE: " & txtjCode, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
            'LogAudit "X", "JOBS", "JCODE" & txtJCode
            ShowDeletedMsg
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    rsRefresh
    StoreMemVars
    FillGrid
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "JOBS") = False Then Exit Sub

    fraDetails.Enabled = False

    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
    txtjCode.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rsROJOBS.MoveNext
    If rsROJOBS.EOF Then
        rsROJOBS.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsROJOBS.MovePrevious
    If rsROJOBS.BOF Then
        rsROJOBS.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode
    Dim rsfindDup                                  As New ADODB.Recordset
    
    If IsNull(txtjCode.Text) = True Then
        MsgSpeechBox "Job Code must not be empty"
        On Error Resume Next
        txtjCode.SetFocus
        Exit Sub
    End If
    If txtDesc1.Text = "" Then
        MsgSpeechBox "Description is Required"
        On Error Resume Next
        txtDesc1.SetFocus
        Exit Sub
    End If
    
    If AddorEdit = "ADD" Then
        rsfindDup.Open "select jcode from CSMS_JobMast where jcode = '" & txtjCode.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsfindDup.EOF And Not rsfindDup.BOF Then
            MsgBox "Job Code already exist!", vbInformation, "Info"
            On Error Resume Next
            txtjCode.SetFocus
            Exit Sub
        End If
    ElseIf AddorEdit = "EDIT" Then
        rsfindDup.Open "select jcode, ID from CSMS_JobMast where jcode = '" & txtjCode.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not (rsfindDup.BOF And rsfindDup.EOF) Then
            If labid.Caption <> rsfindDup!ID Then
                MsgBox "Job Code already exist!", vbInformation, "Info"
                On Error Resume Next
                txtjCode.SetFocus
                Exit Sub
            End If
        End If
    End If

    Dim VTXTJCode                                       As String
    Dim VtxtDesc1                                       As String
    Dim VTXTStd_mHrs                                    As Double
    Dim VTXTFlatrate                                    As Double
    Dim VTXTPOCode                                      As String
    Dim VTXTValidate                                    As String

    VTXTJCode = N2Str2Null(txtjCode.Text)
    VtxtDesc1 = N2Str2Null(txtDesc1.Text)
    VTXTStd_mHrs = NumericVal(txtStd_mHrs.Text)
    VTXTFlatrate = NumericVal(txtflatrate.Text)

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "Insert into CSMS_JobMast" & _
                      " (MAIN_CAT,jcode,description,DETAIL,std_mhrs,flatrate)" & _
                      " values (" & JOB_GROUP & "," & VTXTJCode & ", " & VtxtDesc1 & ", " & VtxtDesc1 & ", " & VTXTStd_mHrs & ", " & _
                      " " & VTXTFlatrate & ")"
        gconDMIS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("A", "JOBS", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtjCode), "JCODE", "CSMS_JOBMAST"), "", "JOB CODE: " & txtjCode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update CSMS_JobMast set" & _
                      " MAIN_CAT = " & JOB_GROUP & "," & _
                      " jcode = " & VTXTJCode & "," & _
                      " description = " & VtxtDesc1 & "," & _
                      " DETAIL = " & VtxtDesc1 & "," & _
                      " std_mhrs = " & VTXTStd_mHrs & "," & _
                      " flatrate = " & VTXTFlatrate & _
                      " where id = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "JOBS", SQL_STATEMENT, labid, "", "JOB CODE: " & txtjCode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        ShowSuccessFullyUpdated
    End If
    rsRefresh
    On Error Resume Next

    rsROJOBS.Find "ID =  " & labid
    cmdCancel.Value = True
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            If Picture1.Visible = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (JOBS MASTER FILE)"
            Call frmALL_AuditInquiry.DisplayHistory(labid, "JOBS", "")

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    optGENERAL.Value = True
    rsRefresh
    Frame1.Enabled = False
    textSearch.Text = "":
    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCSMSJobs = Nothing
End Sub

Private Sub lstJobs_GotFocus()
    On Error Resume Next
    rsROJOBS.Bookmark = rsFind(rsROJOBS.Clone, "JCode", lstJObs.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lstJObs_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    rsROJOBS.Bookmark = rsFind(rsROJOBS.Clone, "JCode", lstJObs.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lstJobs_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstJObs
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstJObs_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstJobs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub optBODY_Click()
    RefreshRecords
End Sub

Private Sub optCHASSIS_Click()
    RefreshRecords
End Sub

Private Sub optELECTRICAL_Click()
    RefreshRecords
End Sub

Private Sub optENGINE_Click()
    RefreshRecords
End Sub

Private Sub optFUEL_Click()
    RefreshRecords
End Sub

Private Sub optGENERAL_Click()
    RefreshRecords
End Sub

Private Sub optPAINT_Click()
    RefreshRecords
End Sub

Private Sub optPMS_Click()
    If COMPANY_CODE = "HGC" Then
        MsgBox "This Module is been disabled", vbInformation, "CSMS"
        Exit Sub
    End If
    RefreshRecords
End Sub

Private Sub optSUBLET_Click()
    If COMPANY_CODE = "HGC" Then
        MsgBox "This Module is been disabled", vbInformation, "CSMS"
        Exit Sub
    End If

    RefreshRecords
End Sub

Private Sub optTRANSMISSION_Click()
    RefreshRecords
End Sub

Private Sub optTRIM_Click()
    RefreshRecords
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (textSearch.Text)
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstJObs.Enabled = True And lstJObs.ListItems.Count > 0 Then
            lstJObs.SetFocus
        End If
    End If
End Sub


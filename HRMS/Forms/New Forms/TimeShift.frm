VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmHRMS_TimeShift 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Shift Code"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   Icon            =   "TimeShift.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   5790
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   60
      ScaleHeight     =   855
      ScaleWidth      =   5640
      TabIndex        =   11
      Top             =   4470
      Width           =   5640
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
         Left            =   4860
         MouseIcon       =   "TimeShift.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "TimeShift.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
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
         Left            =   4170
         MouseIcon       =   "TimeShift.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "TimeShift.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
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
         Left            =   3480
         MouseIcon       =   "TimeShift.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "TimeShift.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
         Width           =   705
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
         Left            =   2790
         MouseIcon       =   "TimeShift.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "TimeShift.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
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
         Left            =   2100
         MouseIcon       =   "TimeShift.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "TimeShift.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
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
         Left            =   1410
         MouseIcon       =   "TimeShift.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "TimeShift.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   705
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
         Left            =   720
         MouseIcon       =   "TimeShift.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "TimeShift.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Move to Next Record"
         Top             =   30
         Width           =   705
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
         Left            =   30
         MouseIcon       =   "TimeShift.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "TimeShift.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   60
      Picture         =   "TimeShift.frx":2D71
      ScaleHeight     =   4365
      ScaleWidth      =   1815
      TabIndex        =   6
      Top             =   30
      Width           =   1845
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   30
         MaxLength       =   35
         TabIndex        =   7
         Top             =   60
         Width           =   1755
      End
      Begin MSComctlLib.ListView lsTimeShift 
         Height          =   3855
         Left            =   30
         TabIndex        =   8
         Top             =   480
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   6800
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
         MouseIcon       =   "TimeShift.frx":5AAD
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SHIFT CODE"
            Object.Width           =   3528
         EndProperty
         Picture         =   "TimeShift.frx":5C0F
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   1920
      TabIndex        =   0
      Top             =   210
      Width           =   3765
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   3420
         Width           =   195
      End
      Begin VB.ComboBox cboND 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   420
         TabIndex        =   32
         Top             =   3720
         Width           =   1395
      End
      Begin VB.ComboBox cboOT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   420
         TabIndex        =   31
         Top             =   3330
         Width           =   1395
      End
      Begin VB.CheckBox Check1 
         Caption         =   "include log-in lunch and log-out lunch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox txtGracePeriod 
         Height          =   315
         Left            =   360
         TabIndex        =   26
         Top             =   2520
         Width           =   495
      End
      Begin VB.ComboBox cboLunchIn 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         TabIndex        =   24
         Top             =   2040
         Width           =   1605
      End
      Begin VB.ComboBox cboLunchOut 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1950
         TabIndex        =   23
         Top             =   1290
         Width           =   1605
      End
      Begin VB.ComboBox cboTo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1950
         TabIndex        =   10
         Top             =   2070
         Width           =   1605
      End
      Begin VB.ComboBox cboFrom1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         TabIndex        =   9
         Top             =   1290
         Width           =   1605
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   150
         MaxLength       =   15
         TabIndex        =   2
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "OverTime Start"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2130
         TabIndex        =   30
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Night Differential Start"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1860
         TabIndex        =   29
         Top             =   3780
         Width           =   1785
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Grace period in minutes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1080
         TabIndex        =   25
         Top             =   2640
         Width           =   1980
      End
      Begin VB.Label LABID 
         Caption         =   "**ID**"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2430
         TabIndex        =   5
         Top             =   2010
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1950
         TabIndex        =   4
         Top             =   960
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   3
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Time Shift Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   1
         Top             =   210
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4260
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   20
      Top             =   4470
      Width           =   1440
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
         Left            =   720
         MouseIcon       =   "TimeShift.frx":1997C
         MousePointer    =   99  'Custom
         Picture         =   "TimeShift.frx":19ACE
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
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
         Left            =   30
         MouseIcon       =   "TimeShift.frx":19E0C
         MousePointer    =   99  'Custom
         Picture         =   "TimeShift.frx":19F5E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   345
      Left            =   1920
      TabIndex        =   28
      Top             =   30
      Width           =   3765
      _Version        =   655364
      _ExtentX        =   6641
      _ExtentY        =   609
      _StockProps     =   14
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
End
Attribute VB_Name = "frmHRMS_TimeShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTimeShift                                                       As ADODB.Recordset

Sub rsrefresh()
    Set rsTimeShift = New ADODB.Recordset
    Set rsTimeShift = gconDMIS.Execute("Select * from HRMS_Time_Shift_Code")
End Sub

Sub InitMemvars()
    txtCode = ""
    cboFrom1 = ""
    cboTo1 = ""
    cboLunchOut = ""
    cboLunchIn = ""
    txtGracePeriod = ""
    Check1.Value = 0
    cboND = ""
    cboOT = ""
    Check2.Value = 0
End Sub

Sub StoreMemVars()
    If rsTimeShift.RecordCount = 0 Then
        ShowNoRecord
        txtCode = ""
        cboFrom1 = ""
        cboTo1 = ""
        cboLunchOut = ""
        cboLunchIn = ""
        txtGracePeriod = ""
        Check1.Value = 0
        LABID = ""
        cboOT = ""
        cboND = ""
        Check2.Value = 0
    End If
    If Not rsTimeShift.BOF And Not rsTimeShift.EOF Then
        txtCode = Null2String(rsTimeShift!shiftcode)
        cboFrom1 = Format(rsTimeShift!FROM1, "HH:MM AM/PM")
        cboTo1 = Format(rsTimeShift!TO1, "HH:MM AM/PM")
        cboLunchOut = Format(rsTimeShift!LUNCHOUT, "HH:MM AM/PM")
        cboLunchIn = Format(rsTimeShift!LUNCHIN, "HH:MM AM/PM")
        txtGracePeriod = Null2String(rsTimeShift!GRACE_PERIOD)
        Check1.Value = N2Str2Zero(rsTimeShift!PATS)
        LABID = Null2String(rsTimeShift!ID)
        cboND = Format(rsTimeShift!NDSTART, "HH:MM AM/PM")
        cboOT = Format(rsTimeShift!OTSTART, "HH:MM AM/PM")
        'Check2.Value = N2Str2Zero(rsTimeShift!OTCOMPUTE)
        Check2.Value = N2Str2Zero(rsTimeShift!OTCOMPUTE)
    End If
End Sub

Sub FillGrid()
    Listview_Loadval lsTimeShift.ListItems, gconDMIS.Execute("select ShiftCode from HRMS_Time_Shift_Code")
End Sub

Sub FillCombo()
    Dim X                                                             As Integer
    X = 0
    Time = #6:00:00 AM#
    Do While Not X = 48
        cboFrom1.AddItem Time
        cboTo1.AddItem Time
        cboLunchOut.AddItem Time
        cboLunchIn.AddItem Time
        cboOT.AddItem Time
        cboND.AddItem Time
        Time = DateAdd("n", 30, Time)
        X = X + 1
    Loop
End Sub

Private Sub cmdAdd_Click()
    Frame1.Caption = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    InitMemvars
End Sub

Private Sub cmdCancel_Click()
    Picture1.Visible = True
    Frame1.Enabled = False
End Sub

Private Sub cmdDelete_Click()
    If LABID <> "" Then
        If ShowConfirmDelete = True Then
            gconDMIS.Execute "delete from HRMS_Time_Shift_Code where ID = " & LABID
        End If
        rsrefresh
        StoreMemVars
        FillGrid
    End If
End Sub

Private Sub cmdEdit_Click()
    Frame1.Caption = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    txtsearch.SetFocus
    txtsearch = ""
End Sub

Private Sub cmdNext_Click()
    If Not rsTimeShift.EOF And Not rsTimeShift.EOF Then
        rsTimeShift.MoveNext
        If rsTimeShift.EOF Then
            rsTimeShift.MoveLast
        End If
        StoreMemVars
    End If
End Sub

Private Sub cmdPrevious_Click()
    If Not rsTimeShift.EOF And Not rsTimeShift.EOF Then
        rsTimeShift.MovePrevious
        If rsTimeShift.BOF Then
            rsTimeShift.MoveFirst
        End If
        StoreMemVars
    End If
End Sub

Private Sub cmdSave_Click()
    Dim vtxtCode                                                      As String
    Dim vcboFrom1                                                     As String
    Dim vcboTo1                                                       As String
    Dim vcboLunchOut                                                  As String
    Dim vcboLunchIn                                                   As String
    Dim vtxtGracePeriod                                               As String
    Dim vtxtPATS                                                      As Integer
    Dim vcboOT                                                        As String
    Dim vcboND                                                        As String
    Dim vtxtOTCompute                                                 As Integer

    vtxtCode = N2Str2Null(txtCode)
    vcboFrom1 = N2Str2Null(cboFrom1)
    vcboTo1 = N2Str2Null(cboTo1)
    vcboLunchOut = N2Str2Null(cboLunchOut)
    vcboLunchIn = N2Str2Null(cboLunchIn)
    vtxtGracePeriod = N2Str2Null(txtGracePeriod)
    vtxtPATS = Check1.Value
    vcboOT = N2Str2Null(cboOT)
    vcboND = N2Str2Null(cboND)
    vtxtOTCompute = Check2.Value

    If Frame1.Caption = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_Time_Shift_Code (ShiftCode, From1, To1, LunchOut, lunchin, Grace_Period, PATS, OTSTART, NDSTART, OTCOMPUTE)" & _
                       " values (" & vtxtCode & _
                         "," & vcboFrom1 & _
                         "," & vcboTo1 & _
                         "," & vcboLunchOut & _
                         "," & vcboLunchIn & _
                         "," & vtxtGracePeriod & _
                         "," & vtxtPATS & _
                         "," & vcboOT & _
                         "," & vcboND & _
                         "," & vtxtOTCompute & ")"
    Else
        If LABID <> "" Then
            gconDMIS.Execute "Update HRMS_Time_Shift_Code set" & _
                           " ShiftCode = " & vtxtCode & "," & _
                           " from1 = " & vcboFrom1 & "," & _
                           " To1 = " & vcboTo1 & "," & _
                           " LunchOut = " & vcboLunchOut & "," & _
                           " LunchIn = " & vcboLunchIn & "," & _
                           " Grace_Period = " & vtxtGracePeriod & "," & _
                           " PATS = " & vtxtPATS & ", " & _
                           " OTSTART = " & vcboOT & ", " & _
                           " NDSTART = " & vcboND & ", " & _
                           " OTCOMPUTE = " & vtxtOTCompute & _
                           " where ID = " & LABID
        Else
            ShowNoRecord
        End If
    End If
    rsrefresh
    '    rsTimeShift.Find "ShiftCode = " & txtCode
    StoreMemVars
    FillGrid
    cmdCancel.Value = True
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Frame1.Enabled = False
    InitMemvars
    rsrefresh
    StoreMemVars
    FillCombo
    FillGrid
    Screen.MousePointer = 0
End Sub

Private Sub lsTimeShift_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsTimeShift.Bookmark = rsFIND(rsTimeShift.Clone, "ShiftCode", Me.lsTimeShift.SelectedItem).Bookmark
    StoreMemVars
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtsearch_Change()
    Listview_Loadval lsTimeShift.ListItems, gconDMIS.Execute("select ShiftCode from HRMS_Time_Shift_Code where ShiftCode like '%" & Repleys(txtsearch) & "%'")
End Sub


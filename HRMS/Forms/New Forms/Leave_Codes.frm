VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmHRMS_Leave_Codes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leave Codes"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9180
   Icon            =   "Leave_Codes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   9180
   Begin VB.PictureBox FRame1 
      Height          =   2595
      Left            =   3570
      ScaleHeight     =   2535
      ScaleWidth      =   5385
      TabIndex        =   15
      Top             =   60
      Width           =   5445
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3930
         TabIndex        =   24
         Top             =   2040
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "# of leaves fixed in a year"
         Height          =   345
         Left            =   210
         TabIndex        =   23
         Top             =   1680
         Width           =   2985
      End
      Begin VB.OptionButton Option1 
         Caption         =   "# of leaves accumulated every month"
         Height          =   255
         Left            =   210
         TabIndex        =   22
         Top             =   1410
         Width           =   3045
      End
      Begin VB.TextBox txtLeave_Code 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2220
         TabIndex        =   17
         Top             =   510
         Width           =   3135
      End
      Begin VB.TextBox txtLeave_Desc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2220
         TabIndex        =   16
         Top             =   930
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "Input number of leaves here in a year/month >>"
         Height          =   225
         Left            =   300
         TabIndex        =   25
         Top             =   2100
         Width           =   3525
      End
      Begin VB.Label LABID 
         Height          =   345
         Left            =   1620
         TabIndex        =   21
         Top             =   510
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Leave Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   20
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Leave Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   19
         Top             =   960
         Width           =   1575
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   345
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   5505
         _Version        =   655364
         _ExtentX        =   9710
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
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3540
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   6
      Top             =   2700
      Width           =   5580
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
         MouseIcon       =   "Leave_Codes.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Codes.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   14
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
         MouseIcon       =   "Leave_Codes.frx":0A42
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Codes.frx":0B94
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
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
         MouseIcon       =   "Leave_Codes.frx":0EFA
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Codes.frx":104C
         Style           =   1  'Graphical
         TabIndex        =   12
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
         MouseIcon       =   "Leave_Codes.frx":1377
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Codes.frx":14C9
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
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
         MouseIcon       =   "Leave_Codes.frx":1825
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Codes.frx":1977
         Style           =   1  'Graphical
         TabIndex        =   10
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
         MouseIcon       =   "Leave_Codes.frx":1C8A
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Codes.frx":1DDC
         Style           =   1  'Graphical
         TabIndex        =   9
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
         MouseIcon       =   "Leave_Codes.frx":20D6
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Codes.frx":2228
         Style           =   1  'Graphical
         TabIndex        =   8
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
         MouseIcon       =   "Leave_Codes.frx":2580
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Codes.frx":26D2
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7650
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   3
      Top             =   2670
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
         MouseIcon       =   "Leave_Codes.frx":2A31
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Codes.frx":2B83
         Style           =   1  'Graphical
         TabIndex        =   5
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
         MouseIcon       =   "Leave_Codes.frx":2EC1
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Codes.frx":3013
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picSearch 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3600
      Left            =   0
      Picture         =   "Leave_Codes.frx":3363
      ScaleHeight     =   3570
      ScaleWidth      =   3465
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
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
         Left            =   30
         MaxLength       =   35
         TabIndex        =   1
         Top             =   30
         Width           =   3405
      End
      Begin MSComctlLib.ListView lsAdjustment 
         Height          =   3105
         Left            =   30
         TabIndex        =   2
         Top             =   420
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   5477
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
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
         MouseIcon       =   "Leave_Codes.frx":609F
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPTION"
            Object.Width           =   5292
         EndProperty
         Picture         =   "Leave_Codes.frx":6201
      End
   End
End
Attribute VB_Name = "frmHRMS_Leave_Codes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsLeave_Code                                                      As ADODB.Recordset
Dim ADDOREDIT                                                         As String

Function CheckIfExist(CHOICE As String, LEAVE_CODE As String) As Boolean
    Dim rsTemp                                                        As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_LEAVEMASTER WHERE LEAVE_CODE = '" & LEAVE_CODE & "'")
    CheckIfExist = False
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        If CHOICE = "ADD" Then
            CheckIfExist = True
        ElseIf CHOICE = "EDIT" Then
            If rsTemp.RecordCount >= 2 Then
                CheckIfExist = True
            End If
        End If
    End If
    Set rsTemp = Nothing
End Function

Sub InitMemvars()
    txtLeave_Code = ""
    txtLeave_Desc = ""
    Option1.Value = False
    Option2.Value = False
    Text1.Text = ""
End Sub

Sub StoreMemVars()
    If rsLeave_Code.RecordCount = 0 Then
        ShowNoRecord
        txtLeave_Code = ""
        txtLeave_Desc = ""
        LABID = ""
        Text1.Text = ""
        Option1.Value = False
        Option2.Value = False
    Else
        If Not rsLeave_Code.EOF And Not rsLeave_Code.BOF Then
            txtLeave_Code = Null2String(rsLeave_Code!LEAVE_CODE)
            txtLeave_Desc = Null2String(rsLeave_Code!LEAVE_desc)
            LABID.Caption = Null2String(rsLeave_Code!ID)
            If N2Str2Zero(rsLeave_Code!Type) = 0 Then
                Option2.Value = True
                Option1.Value = False
            Else
                Option1.Value = True
                Option2.Value = False
            End If
            Text1.Text = N2Str2Zero(rsLeave_Code!DAYS_NO)
        End If
    End If
End Sub

Sub rsrefresh()
    Set rsLeave_Code = New ADODB.Recordset
    Set rsLeave_Code = gconDMIS.Execute("SELECT * FROM HRMS_LEAVEMASTER")
End Sub

Sub FillCombo()
    Listview_Loadval lsAdjustment.ListItems, gconDMIS.Execute("SELECT LEAVE_CODE, LEAVE_DESC FROM HRMS_LEAVEMASTER")
End Sub

Private Sub cmdAdd_Click()
    ADDOREDIT = "ADD"
    InitMemvars
    Picture1.Visible = False
    Picture2.Visible = True
    FRame1.Enabled = True
    picSearch.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    Picture1.Visible = True
    Picture2.Visible = False
    FRame1.Enabled = False
    picSearch.Enabled = True
End Sub

Private Sub cmdDelete_Click()
    gconDMIS.Execute ("DELETE FROM HRMS_LEAVEMASTER WHERE ID = '" & LABID & "'")
    ShowDeletedMsg
    rsrefresh
    FillCombo
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    ADDOREDIT = "EDIT"
    Picture1.Visible = False
    Picture2.Visible = True
    FRame1.Enabled = True
    picSearch.Enabled = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    If Not rsLeave_Code.EOF And Not rsLeave_Code.EOF Then
        rsLeave_Code.MoveNext
        If rsLeave_Code.EOF Then
            rsLeave_Code.MoveLast
        End If
        StoreMemVars
    End If
End Sub

Private Sub cmdPrevious_Click()
    If Not rsLeave_Code.EOF And Not rsLeave_Code.EOF Then
        rsLeave_Code.MovePrevious
        If rsLeave_Code.BOF Then
            rsLeave_Code.MoveFirst
        End If
        StoreMemVars
    End If
End Sub

Private Sub cmdSave_Click()
    Dim vtxtLeave_Code                                                As String
    Dim vtxtLeave_Desc                                                As String
    Dim vtxtLeave_Number                                              As Double
    Dim vtxtLeave_Type                                                As Integer

    vtxtLeave_Number = N2Str2Zero(Text1.Text)
    vtxtLeave_Code = N2Str2Null(txtLeave_Code)
    vtxxtLeave_Desc = N2Str2Null(txtLeave_Desc)

    If Option1.Value = True Then
        vtxtLeave_Type = 1
    Else
        vtxtLeave_Type = 0
    End If


    If ADDOREDIT = "ADD" Then
        If CheckIfExist("ADD", txtLeave_Code) = True Then
            ShowAlreadyExistMsg "Code"
        Else
            gconDMIS.Execute "INSERT INTO HRMS_LEAVEMASTER (LEAVE_CODE, LEAVE_DESC, DAYS_NO, TYPE)" & _
                           " VALUES (" & vtxtLeave_Code & "," & vtxxtLeave_Desc & ", " & vtxtLeave_Number & ", " & vtxtLeave_Type & ")"
            ShowSuccessFullyAdded
        End If
    Else
        If CheckIfExist("EDIT", txtLeave_Code) = True Then
            ShowAlreadyExistMsg "Code"
        Else
            gconDMIS.Execute "UPDATE HRMS_LEAVEMASTER SET" & _
                           " LEAVE_CODE = " & vtxtLeave_Code & "," & _
                           " LEAVE_DESC = " & vtxxtLeave_Desc & "," & _
                           " DAYS_NO = " & vtxtLeave_Number & "," & _
                           " TYPE = " & vtxtLeave_Type & _
                           " WHERE ID = '" & LABID & "'"
            ShowSuccessFullyUpdated
        End If
    End If

    rsrefresh
    StoreMemVars
    FillCombo
    ADDOREDIT = ""
    cmdCancel.Value = True
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    'DrawXPCtl Me
    FRame1.Enabled = False
    Picture2.Visible = False
    rsrefresh
    FillCombo
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub lsAdjustment_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    On Error Resume Next
    rsLeave_Code.Bookmark = rsFIND(rsLeave_Code.Clone, "LEAVE_CODE", Me.lsAdjustment.SelectedItem).Bookmark
    StoreMemVars
End Sub

Private Sub txtsearch_Change()
    Listview_Loadval lsAdjustment.ListItems, gconDMIS.Execute("SELECT LEAVE_CODE, LEAVE_DESC FROM HRMS_LEAVEMASTER WHERE LEAVE_CODE LIKE '%" & Repleys(txtSearch) & "%'")
End Sub


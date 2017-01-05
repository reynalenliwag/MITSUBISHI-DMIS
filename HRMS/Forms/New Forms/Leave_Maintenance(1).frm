VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Object = "{A9046457-E246-455F-A58F-D670C44E8BEA}#2.0#0"; "wizFlexCracker.ocx"
Begin VB.Form frmHRMS_Leave_Maintenance_OLD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leave Maintenance"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12360
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Leave_Maintenance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   12360
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4155
      Left            =   13020
      ScaleHeight     =   4125
      ScaleWidth      =   7875
      TabIndex        =   10
      Top             =   2430
      Width           =   7905
      Begin MSComctlLib.ListView lsvDet 
         Height          =   3675
         Left            =   60
         TabIndex        =   17
         Top             =   360
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   6482
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Leave Type"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Avialable"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Used"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Remaining"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   0
         EndProperty
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   7905
         _Version        =   655364
         _ExtentX        =   13944
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "DETAILS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12582912
         GradientColorDark=   8388608
      End
   End
   Begin VB.PictureBox Picture 
      Height          =   1185
      Left            =   12480
      ScaleHeight     =   1125
      ScaleWidth      =   8475
      TabIndex        =   28
      Top             =   810
      Width           =   8535
      Begin VB.TextBox Text1 
         Height          =   345
         Left            =   930
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   360
         Width           =   2985
      End
      Begin VB.TextBox Text2 
         Height          =   345
         Left            =   4770
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   330
         TabIndex        =   32
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4050
         TabIndex        =   31
         Top             =   420
         Width           =   675
      End
   End
   Begin VB.PictureBox picAddDet 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   4590
      ScaleHeight     =   2505
      ScaleWidth      =   3165
      TabIndex        =   18
      Top             =   2663
      Visible         =   0   'False
      Width           =   3195
      Begin VB.TextBox txtUsed 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1020
         TabIndex        =   27
         Top             =   1170
         Width           =   1065
      End
      Begin VB.TextBox txtAvail 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   810
         Width           =   1065
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   2400
         MouseIcon       =   "Leave_Maintenance.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cancel"
         Top             =   1650
         Width           =   705
      End
      Begin VB.ComboBox cboType 
         Height          =   330
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   450
         Width           =   1395
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Save"
         Height          =   795
         Left            =   1710
         MouseIcon       =   "Leave_Maintenance.frx":0A1A
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":0B6C
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Save Entry"
         Top             =   1650
         Width           =   705
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Used"
         Height          =   210
         Index           =   2
         Left            =   570
         TabIndex        =   25
         Top             =   1230
         Width           =   375
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Available"
         Height          =   210
         Index           =   1
         Left            =   285
         TabIndex        =   24
         Top             =   900
         Width           =   660
      End
      Begin VB.Label lblcap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Type"
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   540
         Width           =   855
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   7905
         _Version        =   655364
         _ExtentX        =   13944
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "ADD / EDIT"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   8421504
         GradientColorDark=   4210752
      End
   End
   Begin wizFlexCracker.wizFlexCrack wizFlexCrack1 
      Height          =   3765
      Left            =   3600
      TabIndex        =   13
      Top             =   8040
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   6641
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   60
      ScaleHeight     =   3000
      ScaleWidth      =   12210
      TabIndex        =   7
      Top             =   120
      Width           =   12240
      Begin VB.ComboBox cboyear 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10830
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   30
         Width           =   1275
      End
      Begin VB.TextBox txtSearch 
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
         Height          =   285
         Left            =   1650
         MaxLength       =   35
         TabIndex        =   8
         Top             =   60
         Width           =   4365
      End
      Begin MSComctlLib.ListView lsAdjustment 
         Height          =   2475
         Left            =   0
         TabIndex        =   9
         Top             =   510
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   4366
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
         MouseIcon       =   "Leave_Maintenance.frx":0EBC
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "NAME"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "VL"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "SL"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "EL"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "PL"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "ML"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "EMPNO"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "YEAR"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   10290
         TabIndex        =   35
         Top             =   120
         Width           =   435
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
         Height          =   465
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   12435
         _Version        =   655364
         _ExtentX        =   21934
         _ExtentY        =   820
         _StockProps     =   14
         Caption         =   "SEARCH EMPLOYEE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   8421504
         GradientColorDark=   4210752
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   6720
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   0
      Top             =   6990
      Width           =   5580
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   3450
         MouseIcon       =   "Leave_Maintenance.frx":101E
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":1170
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   4860
         MouseIcon       =   "Leave_Maintenance.frx":14CC
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":161E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   795
         Left            =   4170
         MouseIcon       =   "Leave_Maintenance.frx":1984
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":1AD6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   2760
         MouseIcon       =   "Leave_Maintenance.frx":1E3C
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":1F8E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Add Record"
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   10890
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   4
      Top             =   6990
      Visible         =   0   'False
      Width           =   1440
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   720
         MouseIcon       =   "Leave_Maintenance.frx":22A1
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":23F3
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   30
         MouseIcon       =   "Leave_Maintenance.frx":2731
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":2883
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   3855
      Left            =   60
      ScaleHeight     =   3795
      ScaleWidth      =   12195
      TabIndex        =   11
      Top             =   3150
      Width           =   12255
      Begin XtremeReportControl.ReportControl rptLIST 
         Height          =   3345
         Left            =   30
         TabIndex        =   36
         Top             =   360
         Width           =   12105
         _Version        =   655364
         _ExtentX        =   21352
         _ExtentY        =   5900
         _StockProps     =   64
      End
      Begin wizButton.cmd cmdX 
         Height          =   315
         Left            =   12330
         TabIndex        =   15
         Top             =   60
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         TX              =   "X"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "Leave_Maintenance.frx":2BD3
      End
      Begin FlexCell.Grid Grid1 
         Height          =   3375
         Left            =   30
         TabIndex        =   12
         Top             =   4140
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   5953
         BackColorBkg    =   -2147483645
         Cols            =   5
         DefaultFontSize =   8.25
         DisplayRowIndex =   -1  'True
         Rows            =   30
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   315
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   12225
         _Version        =   655364
         _ExtentX        =   21564
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "DETAILS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.26
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12582912
         GradientColorDark=   8388608
      End
   End
   Begin VB.Menu mneMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnEdit1 
         Caption         =   "Edit VL"
      End
      Begin VB.Menu mnEdit2 
         Caption         =   "Edit SL"
      End
      Begin VB.Menu mnEdit3 
         Caption         =   "Edit EL"
      End
      Begin VB.Menu mnEdit4 
         Caption         =   "Edit ML"
      End
      Begin VB.Menu mnEdit5 
         Caption         =   "Edit PL"
      End
   End
End
Attribute VB_Name = "frmHRMS_Leave_Maintenance_OLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmpInfo       As ADODB.Recordset
Dim EL_NO           As Double
Dim ML_NO           As Double
Dim PL_NO           As Double
Dim VL_NO           As Double
Dim SL_NO           As Double
Dim ADDOREDIT       As String
Dim DET_ID          As Integer

Private Sub cboType_Change()
    Call DisplayDefaultValue
End Sub

Sub DisplayDefaultValue()
    Dim RSTMP As New ADODB.Recordset
    Dim rsday As New ADODB.Recordset
    
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_LEAVEMASTER WHERE LEAVE_CODE = " & N2Str2Null(cboType) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        txtAvail = NumericVal(RSTMP!DAYS_NO)
        txtUsed = NumericVal(RSTMP!DAYS_NO)
    End If
'
'     Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_leave WHERE type= " & N2Str2Null(cboType) & "")
'        If Not (RSTMP.BOF And RSTMP.EOF) Then
'        txtAvail = NumericVal(RSTMP!available)
'        txtUsed = NumericVal(RSTMP!used)
'    End If

    Set RSTMP = Nothing
End Sub

Private Sub cboType_Click()
    Call DisplayDefaultValue
End Sub

Private Sub cboType_LostFocus()
    Call DisplayDefaultValue
End Sub

Private Sub cboyear_Click()
    Call StoreMemVars
End Sub

Private Sub cmdAdd_Click()
    ADDOREDIT = "ADD"
    Picture2.Visible = False
    Picture1.Visible = False
    'Enable (True)
    Picture4.Enabled = False
    picSearch.Enabled = False
    cboyear.Enabled = False
    cboType.Enabled = True
    cboType.ListIndex = 0
    picAddDet.Visible = True
    On Error Resume Next
    cboType.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Picture1.Visible = True
    Picture2.Visible = False
    Enable (False)
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    ADDOREDIT = "EDIT"
    picAddDet.Visible = True
    'Picture2.Visible = True
    Picture1.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()

End Sub


Private Sub cmdSave_Click()
'    If Text2.Text <> "" Then
'        Dim rsTemp As ADODB.Recordset
'        Set rsTemp = New ADODB.Recordset
'        Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_LEAVE_BALANCE WHERE EMPNO = '" & Text2.Text & "'")
'        If Not rsTemp.EOF And Not rsTemp.BOF Then
'            gconDMIS.Execute ("UPDATE HRMS_LEAVE_BALANCE SET " & _
'                "SL = " & N2Str2Null(Text4.Text) & ", " & _
'                "VL = " & N2Str2Null(Text5.Text) & ", " & _
'                "EL = " & N2Str2Null(Text6.Text) & ", " & _
'                "ML = " & N2Str2Null(Text7.Text) & ", " & _
'                "PL = " & N2Str2Null(Text8.Text) & ", " & _
'                " WHERE EMPNO = '" & Text2.Text & "' AND YEAR_BALANCE = '" & NumericVal(cboyear.Text) - 1 & "'")
'        Else
'            gconDMIS.Execute ("INSERT INTO HRMS_LEAVE_BALANCE " & _
'                "(EMPNO, SL, VL, EL, ML, PL,YEAR_BALANCE)" & _
'                " VALUES (" & _
'                " " & Text2.Text & _
'                ", " & N2Str2Zero(Text4.Text) & _
'                ", " & N2Str2Zero(Text5.Text) & _
'                ", " & N2Str2Zero(Text6.Text) & _
'                ", " & N2Str2Zero(Text7.Text) & _
'                ", " & N2Str2Zero(Text8.Text) & _
'                ", " & N2Str2Null(cboyear.Text) & ")")
'        End If
'    End If
'
'    Set rsTemp = Nothing
'    cmdCancel.Value = True
End Sub

Private Sub cmdX_Click()
    Picture5.Visible = False
    Picture4.Visible = True
End Sub

Private Sub Command1_Click()
    picAddDet.Visible = False
    
    cboyear.Enabled = True
    Picture1.Visible = True
    Picture4.Enabled = True
    picSearch.Enabled = True
End Sub

Private Sub Command2_Click()
    Dim XTYPE As String
    Dim XEMPNO As String
    Dim XAVAIL As Integer
    Dim XUSED As Integer
    Dim RSTMP As New ADODB.Recordset
    
   
    XTYPE = N2Str2Null(cboType)
    XEMPNO = N2Str2Null(Text2)
    XAVAIL = NumericVal(txtAvail)
    XUSED = NumericVal(txtUsed)
    
    If ADDOREDIT = "ADD" Then
        Set RSTMP = gconDMIS.Execute("SELECT TYPE FROM HRMS_LEAVE WHERE EMPLNO = " & XEMPNO & _
            " AND TYPE = " & XTYPE & " AND YEAR(DATEASOF) = " & cboyear & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            MsgBox "Leave type already exist in this Employee", vbExclamation, "Duplicate Leave Type"
            On Error Resume Next
            cboType.SetFocus
            Exit Sub
        End If
    Else
        
        Set RSTMP = gconDMIS.Execute("SELECT EMPLNO, TYPE FROM HRMS_LEAVE WHERE EMPLNO = " & XEMPNO & _
            " AND TYPE = " & XTYPE & " AND YEAR(DATEASOF) = " & cboyear & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            If Not N2Str2Null(RSTMP!EMPLNO) = XEMPNO Then
                MsgBox "Leave type already exist in this Employee", vbExclamation, "Duplicate Leave Type"
                On Error Resume Next
                cboType.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    If ADDOREDIT = "ADD" Then
        gconDMIS.Execute ("INSERT INTO HRMS_LEAVE (EMPLNO, TYPE, AVAILABLE, USED, DATEASOF) " & _
            " VALUES (" & XEMPNO & _
            ", " & XTYPE & _
            ", " & XAVAIL & _
            ", " & XUSED & _
            ", " & N2Str2Null(Date) & ")")
        
        Call KindOfLeave(XAVAIL)
        
        Call ShowSuccessFullyAdded
    Else
        gconDMIS.Execute ("UPDATE HRMS_LEAVE SET AVAILABLE = " & XAVAIL & _
            ", USED =  " & XUSED & _
            " WHERE TYPE = " & XTYPE & _
            " AND EMPLNO = " & XEMPNO & _
            " AND YEAR(DATEASOF) = " & cboyear & "")
            
        Call ShowSuccessFullyUpdated
    End If
    
    'Call DisplayLeave(Text2)
    Call Command1_Click
End Sub

Function KindOfLeave(xNum As Integer)
 Dim sqltxt As String
     
  If cboType.Text = "VL" Then
    sqltxt = "update hrms_leave set maxVL = '" & Trim(xNum) & "'"
    sqltxt = sqltxt & " where emplno = '" & Text2.Text & "'  and type = 'VL'"
        
        gconDMIS.Execute (sqltxt)
     
  ElseIf cboType.Text = "SL" Then
    sqltxt = "update hrms_leave set maxSL = '" & Trim(xNum) & "'"
        sqltxt = sqltxt & " where emplno = '" & Text2.Text & "' and type = 'SL'"

        gconDMIS.Execute (sqltxt)
  ElseIf cboType.Text = "ML" Then
    sqltxt = "update hrms_leave set maxML = '" & Trim(xNum) & "'"
        sqltxt = sqltxt & " where emplno = '" & Text2.Text & "' and type = 'ML'"

        gconDMIS.Execute (sqltxt)
  ElseIf cboType.Text = "EL" Then
    sqltxt = "update hrms_leave set maxEL = '" & Trim(xNum) & "'"
        sqltxt = sqltxt & " where emplno = '" & Text2.Text & "' and type = 'EL'"

        gconDMIS.Execute (sqltxt)
  ElseIf cboType.Text = "PL" Then
    sqltxt = "update hrms_leave set maxPL = '" & Trim(xNum) & "'"
        sqltxt = sqltxt & " where emplno = '" & Text2.Text & "'and type = 'PL'"

        gconDMIS.Execute (sqltxt)
  End If
  
End Function


Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    DrawXPCtl Me
    Call FillCombo
    Call rsrefresh:     'Call FillGrid:
    Call txtsearch_Change
    Call FillLeaveType
    Call InitMemvars:   Call StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub FillLeaveType()
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT LEAVE_CODE FROM HRMS_LEAVEMASTER ORDER BY LEAVE_CODE")
    cboType.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cboType.AddItem Null2String(RSTMP!LEAVE_CODE)
            
            RSTMP.MoveNext
        Loop
        cboType.ListIndex = 0
    End If
    Set RSTMP = Nothing
End Sub

Sub rsrefresh()
    Set rsEmpInfo = New ADODB.Recordset
    Set rsEmpInfo = gconDMIS.Execute("SELECT LASTNAME + ', ' + FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE <> 'I' ORDER BY LASTNAME + ', ' + FIRSTNAME")
End Sub

Sub FillGrid()
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEmpInfo
    End If
End Sub

Private Sub lsAdjustment_DblClick()
ADDOREDIT = "EDIT"
picAddDet.Visible = True
Picture1.Visible = False

End Sub

Private Sub lsAdjustment_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    Dim Col As Integer
    
    Dim Index As Integer
    Index = lsAdjustment.SelectedItem.Index
    'update by:NVB
    '9/29/2009
    'Dec: get emloyee id
    Text2.Text = ITEM.ListSubItems(6)
    Call DisplaySearch(ITEM.ListSubItems(6))
End Sub

Sub StoreMemVars(Optional Col As Integer)
    On Error Resume Next
    If Text2.Text <> "" Then
        Dim rsTemp As ADODB.Recordset
'        Set rsTemp = New ADODB.Recordset
'        Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_LEAVE_BALANCE WHERE EMPNO = '" & Text2.Text & "' AND YEAR_BALANCE = '" & NumericVal(cboyear.Text) - 1 & "'")
'        If Not rsTemp.EOF And Not rsTemp.BOF Then
'            Text4.Text = N2Str2Zero(rsTemp!SL)
'            Text5.Text = N2Str2Zero(rsTemp!VL)
'            Text6.Text = N2Str2Zero(rsTemp!EL)
'            Text7.Text = N2Str2Zero(rsTemp!ML)
'            Text8.Text = N2Str2Zero(rsTemp!PL)
'        Else
'            InitTextLastYearBalance
'        End If
'        Call ComputeLeaveTaken(Text2.Text, cboyear.Text)
'        Text13.Text = SL_NO
'        Text14.Text = VL_NO
'        Text15.Text = EL_NO
'        Text16.Text = ML_NO
'        Text17.Text = PL_NO
        
        'Call DisplayLeave(Text2.Text)
    Else
        Call InitMemvars
    End If
    Set rsTemp = Nothing
End Sub

Function DisplayLeave(XEMPNO As String, XTYPE As String) As String
    Dim ITEM As ListItem
    Dim RSTMP As New ADODB.Recordset
        
    Grid1.Rows = 1
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_LEAVE WHERE EMPLNO = " & XEMPNO & " AND YEAR(DATEASOF) = " & cboyear & " AND TYPE = " & N2Str2Null(XTYPE) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If RSTMP!MAXVL <> 0 And IsNull(RSTMP!MAXVL) = False Then
            DisplayLeave = "" & (RSTMP!MAXVL - RSTMP!used) & " - " & RSTMP!used
       
        ElseIf RSTMP!MAXSL <> 0 And IsNull(RSTMP!MAXSL) = False Then
            DisplayLeave = "" & (RSTMP!MAXSL - RSTMP!used) & " - " & RSTMP!used
        
        ElseIf RSTMP!MAXPL <> 0 And IsNull(RSTMP!MAXPL) = False Then
            DisplayLeave = "" & (RSTMP!MAXPL - RSTMP!used) & " - " & RSTMP!used
       
        ElseIf RSTMP!MAXML <> 0 And IsNull(RSTMP!MAXML) = False Then
            DisplayLeave = "" & (RSTMP!MAXML - RSTMP!used) & " - " & RSTMP!used
        
        ElseIf RSTMP!MAXEL <> 0 And IsNull(RSTMP!MAXEL) = False Then
            DisplayLeave = "" & (RSTMP!MAXEL - RSTMP!used) & " - " & RSTMP!used
        End If
        
    End If
    Set RSTMP = Nothing
End Function

Function GetLeaveDescription(XLEAVETYPE As String) As String
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT LEAVE_DESC FROM HRMS_LEAVEMASTER WHERE LEAVE_CODE = " & XLEAVETYPE & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetLeaveDescription = Null2String(RSTMP!LEAVE_desc)
    End If
    Set RSTMP = Nothing
End Function

Sub InitMemvars()
    InitTextLastYearBalance
    InitTextThisYearBalance
    InitLeavesTaken
End Sub

Sub FillCombo()
    FillcboYear cboyear
End Sub

Sub InitTextLastYearBalance()
'    Text4.Text = ""
'    Text5.Text = ""
'    Text6.Text = ""
'    Text7.Text = ""
'    Text8.Text = ""
End Sub

Sub InitTextThisYearBalance()
'    Text3.Text = ""
'    Text9.Text = ""
'    Text10.Text = ""
'    Text11.Text = ""
'    Text12.Text = ""
End Sub

Sub InitLeavesTaken()
'    Text13.Text = ""
'    Text14.Text = ""
'    Text15.Text = ""
'    Text16.Text = ""
'    Text17.Text = ""
End Sub

Sub ComputeLeaveTaken(EMPNO As String, LEAVE_YEAR As Integer)
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_LEAVEDET WHERE EMPNO = '" & EMPNO & "' AND YEAR(DATEFROM) = '" & LEAVE_YEAR & "'")
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        rsTemp.MoveFirst
        While Not rsTemp.EOF
            If Null2String(rsTemp!LEAVETYPE) = "SL" Then
                SL_NO = SL_NO + N2Str2Zero(rsTemp!DAYS_NO)
            ElseIf Null2String(rsTemp!LEAVETYPE) = "EL" Then
                EL_NO = EL_NO + N2Str2Zero(rsTemp!DAYS_NO)
            ElseIf Null2String(rsTemp!LEAVETYPE) = "VL" Then
                VL_NO = VL_NO + N2Str2Zero(rsTemp!DAYS_NO)
            ElseIf Null2String(rsTemp!LEAVETYPE) = "ML" Then
                ML_NO = ML_NO + N2Str2Zero(rsTemp!DAYS_NO)
            ElseIf Null2String(rsTemp!LEAVETYPE) = "PL" Then
                PL_NO = PL_NO + N2Str2Zero(rsTemp!DAYS_NO)
            End If
            rsTemp.MoveNext
        Wend
    Else
        InitLeavesTaken
    End If
End Sub

Sub Enable(CONDITION As Boolean)
'    Text4.Text = CONDITION
'    Text5.Text = CONDITION
'    Text6.Text = CONDITION
'    Text7.Text = CONDITION
'    Text8.Text = CONDITION
End Sub

Private Sub lsAdjustment_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lsAdjustment.ListItems.count = 0 Then Exit Sub
    
    If Button = vbRightButton Then
        Dim Index As Integer
        
        Index = lsAdjustment.SelectedItem.Index
        
        
        'DET_ID = lsAdjustment.ListItems(INDEX).ListSubItems(4)
        PopupMenu mneMenu
    End If
End Sub

Private Sub lsvDET_DblClick()
    ADDOREDIT = "EDIT"
    
    picSearch.Enabled = False
    Picture4.Enabled = False
    Picture1.Visible = False
    cboType.Enabled = False
    cboyear.Enabled = False
    
    picAddDet.Visible = True
End Sub

Private Sub lsvDet_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    cboType.Text = ITEM.ListSubItems(5)
    txtAvail.Text = ITEM.ListSubItems(1)
    txtUsed.Text = ITEM.ListSubItems(2)
    
    DET_ID = ITEM.ListSubItems(4)
End Sub

Private Sub lsvDet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lsvDet.ListItems.count = 0 Then Exit Sub
    
    If Button = vbRightButton Then
        Dim Index As Integer
        
        Index = lsvDet.SelectedItem.Index
        
        
        DET_ID = lsvDet.ListItems(Index).ListSubItems(4)
        PopupMenu mneMenu
    End If
End Sub

Private Sub mneDetails_Click()
    Dim RSTMP As New ADODB.Recordset
    Dim Index As Integer
    
    Picture5.Visible = True
    Picture4.Visible = False
    Index = lsvDet.SelectedItem.Index
    
    Grid1.Rows = 1
    Call InitGrid
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_REQUESTLEAVE_OT " & _
        " WHERE STATUS = 'A' AND PAY_YEAR = " & cboyear & _
        " AND EMPNO = " & N2Str2Null(Text2) & _
        " AND REQCODE = " & N2Str2Null(lsvDet.ListItems(Index).SubItems(5)) & _
        " AND REQTYPE = 'L'")
        
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Grid1.AddItem Null2String(RSTMP!REASON_REQ) & Chr(9) & _
                Format(RSTMP!DTE_FROM, "MM/DD/YYYY") & Chr(9) & _
                Format(RSTMP!dte_to, "MM/DD/YYYY") & Chr(9) & _
                DateDiff("d", RSTMP!DTE_FROM, RSTMP!dte_to) + 1 & Chr(9) & _
                Null2String(RSTMP!APPROVEDBY), False
                
            RSTMP.MoveNext
        Loop
        Grid1.Refresh:              Grid1.AutoRedraw = True
    End If
    Set RSTMP = Nothing
End Sub

Sub DisplayDetails(Col As Integer, XEMPNO As String)
    Dim RSTMP As New ADODB.Recordset
    Dim Index As Integer
    Dim XTYPE As String
    Picture5.Visible = True
    Picture4.Visible = False
    Index = lsvDet.SelectedItem.Index
    
    Call DisplaySearch(XEMPNO)
End Sub

Sub InitGrid()
    With Grid1
        .Cols = 6
        .FixedCols = 0
        
        .Cell(0, 0).Text = "L/N"
        .Column(0).Width = 30
        
        .Cell(0, 1).Text = "Remarks"
        .Cell(0, 1).Alignment = cellLeftGeneral
        .Column(1).Width = 150
        
        .Cell(0, 2).Text = "From"
        .Cell(0, 2).Alignment = cellLeftGeneral
        .Column(2).Width = 80
    
        .Cell(0, 3).Text = "To"
        .Cell(0, 3).Alignment = cellLeftGeneral
        .Column(3).Width = 80
    
        .Cell(0, 4).Text = "# of Days"
        .Cell(0, 4).Alignment = cellCenterGeneral
        .Column(4).Width = 60
    
        .Cell(0, 5).Text = "Approved By"
        .Cell(0, 5).Alignment = cellLeftGeneral
        .Column(5).Width = 100
    End With
End Sub

Private Sub mnuDelete_Click()
    If MsgBox("Delete this leave type for this employee, are you sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    
    gconDMIS.Execute ("DELETE FROM HRMS_LEAVE WHERE ID = " & DET_ID & "")
    
    Call ShowDeletedMsg
    'Call DisplayLeave(Text2)
    
End Sub

Private Sub mnEdit1_Click()
'    Dim RSTMP As New ADODB.Recordset
'
'    Set RSTMP = gconDMIS.Execute()
End Sub

Private Sub txtsearch_Change()
    Call FillSearchGrid(txtSearch)
End Sub

Sub FillSearchGrid(XXX As String)
    Dim RSTMP As New ADODB.Recordset
    Dim ITEM As ListItem
    
    If XXX = "" Then
        Set RSTMP = gconDMIS.Execute("SELECT TOP 10 LASTNAME + ', ' + FIRSTNAME AS FULLNAME, EMPNO FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE = 'A' ORDER BY LASTNAME")
    Else
        Set RSTMP = gconDMIS.Execute("SELECT TOP 10 LASTNAME + ', ' + FIRSTNAME AS FULLNAME, EMPNO FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE = 'A' AND LASTNAME + ', ' + FIRSTNAME LIKE '%" & XXX & "%' ORDER BY LASTNAME")
    End If
    
    lsAdjustment.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsAdjustment.ListItems.Add(, , Null2String(RSTMP!FULLNAME))
            ITEM.SubItems(1) = DisplayLeave(Null2String(RSTMP!EMPNO), "VL")
            ITEM.SubItems(2) = DisplayLeave(Null2String(RSTMP!EMPNO), "SL")
            ITEM.SubItems(3) = DisplayLeave(Null2String(RSTMP!EMPNO), "EL")
            ITEM.SubItems(4) = DisplayLeave(Null2String(RSTMP!EMPNO), "PL")
            ITEM.SubItems(5) = DisplayLeave(Null2String(RSTMP!EMPNO), "ML")
            ITEM.SubItems(6) = Null2String(RSTMP!EMPNO)
            
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Sub DisplaySearch(XEMPNO As String)
    Screen.MousePointer = 11
    Call ReportControlAddColumnHeader(rptLIST, "Leave type, Remarks, From, To, # of Days, Approved by")
    Call ReportControlPaintManager(rptLIST)
    Call ResizeColumnHeader(rptLIST, " 5, 35, 10, 10, 10, 15")
    rptLIST.GroupsOrder.Add rptLIST.Columns(0)
    rptLIST.Columns(0).Visible = False
    rptLIST.Columns(2).Alignment = xtpAlignmentCenter
    rptLIST.Columns(3).Alignment = xtpAlignmentCenter
    rptLIST.Columns(4).Alignment = xtpAlignmentCenter
    Call flex_FillReportView(gconDMIS.Execute("SELECT REQCODE, REASON_REQ, DTE_FROM, DTE_TO, DATEDIFF(d,DTE_FROM, DTE_TO) + 1 AS #OFDAYS, APPROVEDBY FROM HRMS_REQUESTLEAVE_OT " & _
        " WHERE STATUS = 'A' AND PAY_YEAR = " & cboyear & _
        " AND EMPNO = " & N2Str2Null(XEMPNO) & _
        " AND REQTYPE = 'L'"), rptLIST)
    Screen.MousePointer = 0
End Sub

Sub ReportControlAddColumnHeader(lst As ReportControl, StringHeaders As String)
    Dim ar()                                        As String
    Dim I                                           As Integer

    ar = Split(StringHeaders, ",")
    lst.Columns.DeleteAll
    For I = LBound(ar) To UBound(ar)
        lst.Columns.Add I, ar(I), 100, True
    Next
    Erase ar
    StringHeaders = vbNullString
End Sub

Sub ReportControlPaintManager(lst As ReportControl)
    With lst
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.HighlightBackColor = RGB(34, 133, 13)
        .PaintManager.ShadeSortColor = RGB(250, 251, 189)
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
        .PaintManager.GroupRowTextBold = True
        .PaintManager.GroupForeColor = vbBlue
        .PaintManager.ColumnStyle = xtpColumnExplorer
    End With
End Sub

Public Sub ResizeColumnHeader(grd As Object, SizeArray As String)
    grd.Visible = False

    Dim ar()                                        As String
    Dim cWidth                                      As Long
    Dim I                                           As Integer
    Dim scwidth                                     As Long
    ar = Split(SizeArray, ",")
    cWidth = grd.Width

    If TypeOf grd Is ListView Then
        For I = LBound(ar) To UBound(ar)
            If I <= grd.ColumnHeaders.count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.ColumnHeaders(I + 1).Width = scwidth
            End If
        Next
    ElseIf TypeOf grd Is ReportControl Then
        For I = LBound(ar) To UBound(ar)
            If I < grd.Columns.count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.Columns(I).Width = scwidth
            End If
        Next

    End If

    Erase ar
    grd.Visible = True
End Sub

Public Function flex_FillReportView(RS As ADODB.Recordset, grd As XtremeReportControl.ReportControl, Optional ByVal WithSN As Boolean = False)
    Dim fld                                         As ADODB.Field
    Dim j                                           As Long
    Dim REC                                         As XtremeReportControl.ReportRecord

    grd.Records.DeleteAll

    While Not RS.EOF
        j = j + 1

        Set REC = grd.Records.Add
        If WithSN = True Then
            REC.AddItem j
        End If
        For Each fld In RS.FIELDS
            REC.AddItem (Trim(fld.Value))
        Next
        RS.MoveNext
    Wend
    grd.Populate
    Set fld = Nothing
    Set REC = Nothing
    Set RS = Nothing
End Function

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Object = "{A9046457-E246-455F-A58F-D670C44E8BEA}#2.0#0"; "wizFlexCracker.ocx"
Begin VB.Form frmHRMS_Leave_Maintenance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leave Maintenance"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11160
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
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   11160
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   5550
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   0
      Top             =   4830
      Width           =   5580
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   4860
         MouseIcon       =   "Leave_Maintenance.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":06DC
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
         MouseIcon       =   "Leave_Maintenance.frx":0A42
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":0B94
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   795
         Left            =   3480
         MouseIcon       =   "Leave_Maintenance.frx":0EFA
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":104C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   2790
         MouseIcon       =   "Leave_Maintenance.frx":1377
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":14C9
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   2100
         MouseIcon       =   "Leave_Maintenance.frx":1825
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":1977
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "View Leaves"
         Height          =   795
         Left            =   1410
         MouseIcon       =   "Leave_Maintenance.frx":1C8A
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":1DDC
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   720
         MouseIcon       =   "Leave_Maintenance.frx":20D6
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":2228
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Move to Next Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   30
         MouseIcon       =   "Leave_Maintenance.frx":2580
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":26D2
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin wizFlexCracker.wizFlexCrack wizFlexCrack1 
      Height          =   3765
      Left            =   3600
      TabIndex        =   43
      Top             =   8040
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   6641
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   345
      Left            =   7800
      TabIndex        =   31
      Top             =   90
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   345
      Left            =   3960
      TabIndex        =   30
      Top             =   90
      Width           =   2985
   End
   Begin VB.ComboBox cboyear 
      Height          =   330
      Left            =   9690
      TabIndex        =   29
      Top             =   90
      Width           =   1275
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4965
      ScaleWidth      =   3135
      TabIndex        =   12
      Top             =   30
      Width           =   3135
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
         TabIndex        =   13
         Top             =   30
         Width           =   3045
      End
      Begin MSComctlLib.ListView lsAdjustment 
         Height          =   4455
         Left            =   30
         TabIndex        =   14
         Top             =   420
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   7858
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
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Leave_Maintenance.frx":2A31
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "EMPNO"
            Object.Width           =   2
         EndProperty
      End
   End
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
      Height          =   3165
      Left            =   3180
      ScaleHeight     =   3135
      ScaleWidth      =   7875
      TabIndex        =   16
      Top             =   630
      Width           =   7905
      Begin VB.TextBox Text17 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5850
         TabIndex        =   39
         Top             =   2190
         Width           =   615
      End
      Begin VB.TextBox Text16 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5850
         TabIndex        =   38
         Top             =   1980
         Width           =   615
      End
      Begin VB.TextBox Text15 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5850
         TabIndex        =   37
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text14 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5850
         TabIndex        =   36
         Top             =   1140
         Width           =   615
      End
      Begin VB.TextBox Text13 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5850
         TabIndex        =   35
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4380
         TabIndex        =   27
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4380
         TabIndex        =   26
         Top             =   1140
         Width           =   615
      End
      Begin VB.TextBox Text10 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4380
         TabIndex        =   25
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text11 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4380
         TabIndex        =   24
         Top             =   1980
         Width           =   615
      End
      Begin VB.TextBox Text12 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4380
         TabIndex        =   23
         Top             =   2430
         Width           =   615
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2730
         TabIndex        =   21
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2730
         TabIndex        =   20
         Top             =   1140
         Width           =   615
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2730
         TabIndex        =   19
         Top             =   1530
         Width           =   615
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2730
         TabIndex        =   18
         Top             =   1950
         Width           =   615
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2730
         TabIndex        =   17
         Top             =   2400
         Width           =   615
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   285
         Left            =   0
         TabIndex        =   52
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
      Begin VB.Label Label1 
         Caption         =   "Sick Leave"
         Height          =   285
         Left            =   240
         TabIndex        =   49
         Top             =   750
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "Vacation Leave"
         Height          =   285
         Left            =   240
         TabIndex        =   48
         Top             =   1170
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Emergency Leave"
         Height          =   405
         Left            =   240
         TabIndex        =   47
         Top             =   1560
         Width           =   1635
      End
      Begin VB.Label Label4 
         Caption         =   "Maternity Leave"
         Height          =   405
         Left            =   240
         TabIndex        =   46
         Top             =   1980
         Width           =   1365
      End
      Begin VB.Label Label5 
         Caption         =   "Paternity Leave"
         Height          =   405
         Left            =   240
         TabIndex        =   45
         Top             =   2430
         Width           =   1365
      End
      Begin VB.Label Label13 
         Caption         =   "Leave Type"
         Height          =   285
         Left            =   300
         TabIndex        =   44
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label12 
         Caption         =   "Leaves Taken this Year"
         Height          =   225
         Left            =   5460
         TabIndex        =   40
         Top             =   420
         Width           =   1875
      End
      Begin VB.Label Label7 
         Caption         =   "This Year Balance Balance"
         Height          =   225
         Left            =   3960
         TabIndex        =   28
         Top             =   420
         Width           =   1365
      End
      Begin VB.Label Label6 
         Caption         =   "Last Year Balance"
         Height          =   225
         Left            =   2160
         TabIndex        =   22
         Top             =   420
         Width           =   1365
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
      Left            =   9690
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   9
      Top             =   4830
      Width           =   1440
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   720
         MouseIcon       =   "Leave_Maintenance.frx":2B93
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":2CE5
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   30
         MouseIcon       =   "Leave_Maintenance.frx":3023
         MousePointer    =   99  'Custom
         Picture         =   "Leave_Maintenance.frx":3175
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   4185
      Left            =   3180
      ScaleHeight     =   4125
      ScaleWidth      =   7875
      TabIndex        =   41
      Top             =   600
      Visible         =   0   'False
      Width           =   7935
      Begin wizButton.cmd cmdX 
         Height          =   315
         Left            =   7500
         TabIndex        =   51
         Top             =   0
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
         MICON           =   "Leave_Maintenance.frx":34C5
      End
      Begin FlexCell.Grid Grid1 
         Height          =   3675
         Left            =   30
         TabIndex        =   42
         Top             =   390
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   6482
         BackColorBkg    =   -2147483645
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   315
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Width           =   7905
         _Version        =   655364
         _ExtentX        =   13944
         _ExtentY        =   556
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
      Left            =   7080
      TabIndex        =   34
      Top             =   150
      Width           =   675
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
      Left            =   3360
      TabIndex        =   33
      Top             =   150
      Width           =   495
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
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
      Left            =   9150
      TabIndex        =   32
      Top             =   150
      Width           =   495
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   465
      Left            =   3180
      TabIndex        =   15
      Top             =   30
      Width           =   7905
      _Version        =   655364
      _ExtentX        =   13944
      _ExtentY        =   820
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
      GradientColorLight=   8421504
      GradientColorDark=   4210752
   End
End
Attribute VB_Name = "frmHRMS_Leave_Maintenance"
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

Private Sub cboyear_Click()
    StoreMemVars
End Sub

Private Sub cmdAdd_Click()
    Picture2.Visible = True
    Picture1.Visible = False
    Enable (True)
End Sub

Private Sub cmdCancel_Click()
    Picture1.Visible = True
    Picture2.Visible = False
    Enable (False)
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    Picture2.Visible = True
    Picture1.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Picture5.Visible = True
    Picture4.Visible = False
End Sub

Private Sub cmdSave_Click()
    If Text2.Text <> "" Then
        Dim rsTemp As ADODB.Recordset
        Set rsTemp = New ADODB.Recordset
        Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_LEAVE_BALANCE WHERE EMPNO = '" & Text2.Text & "'")
        If Not rsTemp.EOF And Not rsTemp.BOF Then
            gconDMIS.Execute ("UPDATE HRMS_LEAVE_BALANCE SET " & _
                "SL = " & N2Str2Null(Text4.Text) & ", " & _
                "VL = " & N2Str2Null(Text5.Text) & ", " & _
                "EL = " & N2Str2Null(Text6.Text) & ", " & _
                "ML = " & N2Str2Null(Text7.Text) & ", " & _
                "PL = " & N2Str2Null(Text8.Text) & ", " & _
                " WHERE EMPNO = '" & Text2.Text & "' AND YEAR_BALANCE = '" & NumericVal(cboyear.Text) - 1 & "'")
        Else
            gconDMIS.Execute ("INSERT INTO HRMS_LEAVE_BALANCE " & _
                "(EMPNO, SL, VL, EL, ML, PL,YEAR_BALANCE)" & _
                " VALUES (" & _
                " " & Text2.Text & _
                ", " & N2Str2Zero(Text4.Text) & _
                ", " & N2Str2Zero(Text5.Text) & _
                ", " & N2Str2Zero(Text6.Text) & _
                ", " & N2Str2Zero(Text7.Text) & _
                ", " & N2Str2Zero(Text8.Text) & _
                ", " & N2Str2Null(cboyear.Text) & ")")
        End If
    End If
    
    Set rsTemp = Nothing
    cmdCancel.Value = True
End Sub

Private Sub cmdX_Click()
    Picture5.Visible = False
    Picture4.Visible = True
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    DrawXPCtl Me
    FillCombo
    rsRefresh
    FillGrid
    Picture5.Visible = False
    InitMemVars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub rsRefresh()
    Set rsEmpInfo = New ADODB.Recordset
    Set rsEmpInfo = gconDMIS.Execute("SELECT LASTNAME + ', ' + FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE <> 'I' ORDER BY LASTNAME + ', ' + FIRSTNAME")
End Sub

Sub FillGrid()
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEmpInfo
    End If
End Sub

Private Sub lsAdjustment_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    Text1.Text = ITEM.Text
    Text2.Text = ITEM.ListSubItems(1).Text
    StoreMemVars
End Sub

Sub StoreMemVars()
    On Error Resume Next
    If Text2.Text <> "" Then
        Dim rsTemp As ADODB.Recordset
        Set rsTemp = New ADODB.Recordset
        Set rsTemp = gconDMIS.Execute("SELECT * FROM HRMS_LEAVE_BALANCE WHERE EMPNO = '" & Text2.Text & "' AND YEAR_BALANCE = '" & NumericVal(cboyear.Text) - 1 & "'")
        If Not rsTemp.EOF And Not rsTemp.BOF Then
            Text4.Text = N2Str2Zero(rsTemp!SL)
            Text5.Text = N2Str2Zero(rsTemp!VL)
            Text6.Text = N2Str2Zero(rsTemp!EL)
            Text7.Text = N2Str2Zero(rsTemp!ML)
            Text8.Text = N2Str2Zero(rsTemp!PL)
        Else
            InitTextLastYearBalance
        End If
        Call ComputeLeaveTaken(Text2.Text, cboyear.Text)
        Text13.Text = SL_NO
        Text14.Text = VL_NO
        Text15.Text = EL_NO
        Text16.Text = ML_NO
        Text17.Text = PL_NO
    Else
        InitMemVars
    End If
    Set rsTemp = Nothing
End Sub

Sub InitMemVars()
    InitTextLastYearBalance
    InitTextThisYearBalance
    InitLeavesTaken
End Sub

Sub FillCombo()
    FillcboYear cboyear
End Sub

Sub InitTextLastYearBalance()
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
End Sub

Sub InitTextThisYearBalance()
    Text3.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
End Sub

Sub InitLeavesTaken()
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
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
    Text4.Text = CONDITION
    Text5.Text = CONDITION
    Text6.Text = CONDITION
    Text7.Text = CONDITION
    Text8.Text = CONDITION
End Sub

Private Sub txtSearch_Change()
    Call FillSearchGrid(N2Str2Null(txtSearch))
End Sub

Sub FillSearchGrid(XXX As String)
    Dim RSTMP As New ADODB.Recordset
    Dim ITEM As ListItem
    
    If txtSearch.Text = "" Then
        Set RSTMP = gconDMIS.Execute("SELECT LASTNAME + ', ' + FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE = 'A' ORDER BY LASTNAME")
    Else
        Set RSTMP = gconDMIS.Execute("SELECT LASTNAME + ', ' + FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE ACTIVEINACTIVE = 'A' AND LASTNAME + ', ' + FIRSTNAME LIKE " & XXX & " ORDER BY LASTNAME")
    End If
    
    lsAdjustment.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
        
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

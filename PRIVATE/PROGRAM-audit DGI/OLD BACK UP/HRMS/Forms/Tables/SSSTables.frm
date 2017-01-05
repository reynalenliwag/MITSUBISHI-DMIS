VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRMSTables_SSS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SSS Table"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7980
   ForeColor       =   &H00D8E9EC&
   Icon            =   "SSSTables.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   7980
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2835
      ScaleHeight     =   855
      ScaleWidth      =   5625
      TabIndex        =   21
      Top             =   5850
      Width           =   5625
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
         Left            =   4350
         MouseIcon       =   "SSSTables.frx":0442
         MousePointer    =   99  'Custom
         Picture         =   "SSSTables.frx":0594
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Exit Window"
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
         Left            =   3660
         MouseIcon       =   "SSSTables.frx":08FA
         MousePointer    =   99  'Custom
         Picture         =   "SSSTables.frx":0A4C
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Left            =   2970
         MouseIcon       =   "SSSTables.frx":0DA8
         MousePointer    =   99  'Custom
         Picture         =   "SSSTables.frx":0EFA
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   2280
         MouseIcon       =   "SSSTables.frx":120D
         MousePointer    =   99  'Custom
         Picture         =   "SSSTables.frx":135F
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Left            =   1590
         MouseIcon       =   "SSSTables.frx":1659
         MousePointer    =   99  'Custom
         Picture         =   "SSSTables.frx":17AB
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Left            =   900
         MouseIcon       =   "SSSTables.frx":1B03
         MousePointer    =   99  'Custom
         Picture         =   "SSSTables.frx":1C55
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picSSSTable 
      BorderStyle     =   0  'None
      Height          =   2550
      Left            =   0
      ScaleHeight     =   2550
      ScaleWidth      =   7950
      TabIndex        =   7
      Top             =   0
      Width           =   7950
      Begin VB.TextBox txtEmp_SSS 
         Alignment       =   1  'Right Justify
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
         Left            =   3270
         MaxLength       =   100
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2040
         Width           =   1545
      End
      Begin VB.TextBox txtOwner_EC 
         Alignment       =   1  'Right Justify
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
         Left            =   1770
         MaxLength       =   100
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2040
         Width           =   1425
      End
      Begin VB.TextBox txtBracket 
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
         Left            =   1020
         MaxLength       =   8
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   120
         Width           =   885
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
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
         Left            =   1020
         MaxLength       =   30
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   900
         Width           =   1395
      End
      Begin VB.TextBox txtRange1 
         Alignment       =   1  'Right Justify
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
         Left            =   1020
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   510
         Width           =   1395
      End
      Begin VB.TextBox txtRange2 
         Alignment       =   1  'Right Justify
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
         Left            =   3420
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   510
         Width           =   1395
      End
      Begin VB.TextBox txtOwner_SSS 
         Alignment       =   1  'Right Justify
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
         Left            =   180
         MaxLength       =   100
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2040
         Width           =   1545
      End
      Begin Crystal.CrystalReport rptSalaryGrade 
         Left            =   7410
         Top             =   2580
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.Shape Shape3 
         Height          =   2085
         Left            =   5730
         Top             =   150
         Width           =   1905
      End
      Begin VB.Shape Shape2 
         Height          =   1305
         Left            =   90
         Top             =   60
         Width           =   5325
      End
      Begin VB.Shape Shape1 
         Height          =   1155
         Left            =   90
         Top             =   1350
         Width           =   5325
      End
      Begin VB.Image Image1 
         Height          =   570
         Left            =   6300
         Picture         =   "SSSTables.frx":1FB4
         Top             =   1140
         Width           =   750
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SSS Table"
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
         Left            =   5340
         TabIndex        =   20
         Top             =   630
         Width           =   2655
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee's SSS"
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
         Height          =   255
         Left            =   3300
         TabIndex        =   19
         Top             =   1740
         Width           =   1845
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Employer's EC"
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
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   1740
         Width           =   1845
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Contributions"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   1440
         Width           =   1845
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bracket"
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
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   150
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
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
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   960
         Width           =   1725
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Range1"
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
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   540
         Width           =   1725
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Range2"
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
         Height          =   255
         Left            =   2580
         TabIndex        =   11
         Top             =   540
         Width           =   1635
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Employer's SSS"
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
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   1740
         Width           =   1845
      End
      Begin VB.Label labID 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   120
         Width           =   225
      End
      Begin VB.Label labIDprev 
         Caption         =   "IDprev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   90
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   3285
      Left            =   60
      ScaleHeight     =   3285
      ScaleWidth      =   7905
      TabIndex        =   15
      Top             =   2550
      Width           =   7905
      Begin MSComctlLib.ListView lstSSSTable 
         Height          =   3165
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   5583
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
         MouseIcon       =   "SSSTables.frx":3686
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Bracket"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Range From"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Range To"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Salary Credit"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "ER SS"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "ER EC"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "EE SS"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   6435
      ScaleHeight     =   885
      ScaleWidth      =   1755
      TabIndex        =   28
      Top             =   5880
      Width           =   1755
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
         Left            =   765
         MouseIcon       =   "SSSTables.frx":37E8
         MousePointer    =   99  'Custom
         Picture         =   "SSSTables.frx":393A
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Left            =   75
         MouseIcon       =   "SSSTables.frx":3C78
         MousePointer    =   99  'Custom
         Picture         =   "SSSTables.frx":3DCA
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmHRMSTables_SSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSSSTable                                                        As ADODB.Recordset
Dim AddorEdit                                                         As String

Sub rsrefresh()
    Set rsSSSTable = New ADODB.Recordset
    rsSSSTable.Open "select * from HRMS_SSSTable order by Bracket", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Call FillGrid
End Sub

Sub InitMemvars()
    picSSSTable.Enabled = True
    txtBracket.Text = ""
    txtRange1.Text = 0
    txtRange2.Text = 0
    txtCredit.Text = ""
    txtOwner_SSS.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsSSSTable.EOF And Not rsSSSTable.BOF Then
        picSSSTable.Enabled = False
        labID.Caption = rsSSSTable!ID
        txtBracket.Text = Null2String(rsSSSTable!bracket)
        txtRange1.Text = Null2String(rsSSSTable!Range1)
        txtRange2.Text = Null2String(rsSSSTable!Range2)
        txtCredit.Text = Null2String(rsSSSTable!Credit)
        txtOwner_SSS.Text = Null2String(rsSSSTable!Owner_SSS)
        txtOwner_EC.Text = Null2String(rsSSSTable!Owner_EC)
        txtEmp_SSS.Text = Null2String(rsSSSTable!Emp_SSS)
    Else
        Call ShowNoRecord
        If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdAdd.Value = True Else Unload Me
    End If
End Sub

Sub FillGrid()
    Dim rsSSSTable2                                                   As ADODB.Recordset
    lstSSSTable.Sorted = False: lstSSSTable.ListItems.Clear
    lstSSSTable.Enabled = False

    Set rsSSSTable2 = New ADODB.Recordset
    Set rsSSSTable2 = gconDMIS.Execute("select Bracket,Range1,Range2,Credit,Owner_SSS,Owner_EC,Emp_SSS,ID from HRMS_SSSTable")
    If Not (rsSSSTable2.EOF And rsSSSTable2.BOF) Then
        Listview_Loadval Me.lstSSSTable.ListItems, rsSSSTable2
        lstSSSTable.Refresh
        lstSSSTable.Enabled = True
    End If
End Sub

'Upating Code       : AXP-0707200711:51
Private Sub cmdAdd_Click()
    On Error GoTo Errorcode

    If Function_Access(LOGID, "Acess_Add", "TABLE SSS") = False Then Exit Sub
    AddorEdit = "ADD"

    Call InitMemvars
    lstSSSTable.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True

    On Error Resume Next
    txtBracket.SetFocus

    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    picSSSTable.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstSSSTable.Enabled = True
    StoreMemVars
End Sub

'Upating Code       : AXP-0707200711:53
Private Sub cmdDelete_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Delete", "TABLE SSS") = False Then Exit Sub
    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from HRMS_SSSTable where id = " & labID.Caption

        Call LogAudit("X", "DELETE SSS TABLE RECORD", txtBracket.Text)
        Call ShowDeletedMsg
    End If
    Call rsrefresh
    Call StoreMemVars

    Exit Sub

Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200711:53
Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Edit", "TABLE SSS") = False Then Exit Sub
    AddorEdit = "EDIT"
    picSSSTable.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    lstSSSTable.Enabled = False

    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    UnloadForm Me
End Sub

Private Sub cmdFind_Click()
    If lstSSSTable.ListItems.count > 0 And lstSSSTable.Enabled = True Then: lstSSSTable.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsSSSTable.MoveNext
    If rsSSSTable.EOF Then
        rsSSSTable.MoveLast
        Call ShowLastRecordMsg
    End If
    Call StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsSSSTable.MovePrevious
    If rsSSSTable.BOF Then
        rsSSSTable.MoveFirst
        Call ShowFirstRecordMsg
    End If
    Call StoreMemVars
End Sub

'Private Sub cmdPrint_Click()
'Screen.MousePointer = 11
'rptSSSTable.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
'rptSSSTable.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
'rptSSSTable.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
'PrintSQLReport rptSSSTable, HRMS_REPORT_PATH & "SSSTable.rpt", "", DMIS_REPORT_Connection, 1
'Screen.MousePointer = 0
'End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorBracket
    Dim vtxtBracket                                                   As String
    Dim vtxtRange1, vtxtRange2                                        As Double
    Dim vtxtCredit, vtxtOwner_SSS, vtxtOwner_EC, vtxtEmp_SSS

    vtxtBracket = N2Str2Null(txtBracket.Text)
    vtxtRange1 = NumericVal(txtRange1.Text)
    vtxtRange2 = NumericVal(txtRange2.Text)
    vtxtCredit = N2Str2Null(txtCredit.Text)
    vtxtOwner_SSS = N2Str2Null(txtOwner_SSS.Text)
    vtxtOwner_EC = N2Str2Null(txtOwner_EC.Text)
    vtxtEmp_SSS = N2Str2Null(txtEmp_SSS.Text)

    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_SSSTable " & _
                         "(Bracket,Range1,Range2,Credit,Owner_SSS,Owner_EC,Emp_SSS,LastUpdate,USERCODE) " & _
                       " values (" & vtxtBracket & ", " & _
                         "" & vtxtRange1 & ", " & vtxtRange2 & ", " & vtxtCredit & ", " & vtxtOwner_SSS & ", " & vtxtOwner_EC & ", " & vtxtEmp_SSS & ", '" & LOGDATE & "', '" & LOGCODE & "')"

        Call LogAudit("A", "ADD SSS MASTERFILE RECORD", txtBracket.Text)
    Else
        gconDMIS.Execute "update HRMS_SSSTable set" & _
                       " Bracket = " & vtxtBracket & "," & _
                       " Range1 = " & vtxtRange1 & "," & _
                       " Range2 = " & vtxtRange2 & "," & _
                       " Credit = " & vtxtCredit & "," & _
                       " Owner_SSS = " & vtxtOwner_SSS & "," & _
                       " Owner_EC = " & vtxtOwner_EC & "," & _
                       " Emp_SSS = " & vtxtEmp_SSS & "," & _
                       " LastUpdate = '" & LOGDATE & "'," & _
                       " USERCODE = '" & LOGCODE & "'" & _
                       " where id = " & labID.Caption

        Call LogAudit("E", "UPDATE SSS TABLE", txtBracket.Text)
    End If
    Call rsrefresh

    On Error Resume Next
    rsSSSTable.Find "Bracket = " & vtxtBracket
    cmdCancel.Value = True
    Exit Sub

ErrorBracket:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyBracket As Integer, Shift As Integer)
    MoveKeyPress KeyBracket
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    rsrefresh
    StoreMemVars
    FillGrid
    'DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub lstSSSTable_Click()
    '    Dim INDEX As Double
    '
    '    If Not lstSSSTable.ListItems.Count = 0 Then
    '        With lstSSSTable
    '            INDEX = lstSSSTable.SelectedItem.INDEX
    '
    '            txtBracket.Text = .ListItems(INDEX).Text
    '            txtRange1.Text = .ListItems(INDEX).SubItems(1)
    '            txtRange2.Text = .ListItems(INDEX).SubItems(2)
    '            txtCredit.Text = .ListItems(INDEX).SubItems(3)
    '            txtOwner_SSS.Text = .ListItems(INDEX).SubItems(4)
    '            txtOwner_EC.Text = .ListItems(INDEX).SubItems(5)
    '            txtEmp_SSS.Text = .ListItems(INDEX).SubItems(6)
    '
    '        End With
    '    End If
End Sub

Private Sub lstSSSTable_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstSSSTable
        .Sorted = True
        If .SortKey = ColumnHeader.INDEX - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.INDEX - 1
        End If
    End With
End Sub

Private Sub lstSSSTable_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstSSSTable_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    '    On Error Resume Next
    '
    '    rsSSSTable.Bookmark = rsFind(rsSSSTable.Clone, "Bracket", lstSSSTable.ListItems(lstSSSTable.SelectedItem.INDEX).Text).Bookmark
    '    Call StoreMemVars
    Dim INDEX                                                         As Double

    If Not lstSSSTable.ListItems.count = 0 Then
        With lstSSSTable
            INDEX = lstSSSTable.SelectedItem.INDEX

            txtBracket.Text = .ListItems(INDEX).Text
            txtRange1.Text = .ListItems(INDEX).SubItems(1)
            txtRange2.Text = .ListItems(INDEX).SubItems(2)
            txtCredit.Text = .ListItems(INDEX).SubItems(3)
            txtOwner_SSS.Text = .ListItems(INDEX).SubItems(4)
            txtOwner_EC.Text = .ListItems(INDEX).SubItems(5)
            txtEmp_SSS.Text = .ListItems(INDEX).SubItems(6)
            labID.Caption = .ListItems(INDEX).SubItems(7)
        End With
    End If
End Sub


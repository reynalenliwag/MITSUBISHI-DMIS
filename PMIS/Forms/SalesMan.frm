VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISMaster_SalesMan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salesman Master File"
   ClientHeight    =   4005
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   8490
   FillColor       =   &H8000000F&
   ForeColor       =   &H00DEDFDE&
   Icon            =   "SalesMan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4005
   ScaleWidth      =   8490
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   2670
      ScaleHeight     =   885
      ScaleWidth      =   5775
      TabIndex        =   24
      Top             =   3090
      Width           =   5775
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
         Left            =   5040
         MouseIcon       =   "SalesMan.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "SalesMan.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   735
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
         Left            =   4320
         MouseIcon       =   "SalesMan.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "SalesMan.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   735
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
         Left            =   3600
         MouseIcon       =   "SalesMan.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "SalesMan.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Delete Selected Record"
         Top             =   0
         Width           =   735
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
         Left            =   2880
         MouseIcon       =   "SalesMan.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "SalesMan.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   735
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
         Left            =   2160
         MouseIcon       =   "SalesMan.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "SalesMan.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   735
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
         Left            =   1440
         MouseIcon       =   "SalesMan.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "SalesMan.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Find a Record"
         Top             =   0
         Width           =   735
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
         MouseIcon       =   "SalesMan.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "SalesMan.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Move to Next Record"
         Top             =   0
         Width           =   735
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
         Left            =   0
         MouseIcon       =   "SalesMan.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "SalesMan.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   3060
      TabIndex        =   7
      Top             =   60
      Width           =   5355
      Begin VB.TextBox txtPositions 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1380
         TabIndex        =   6
         Text            =   "Text1"
         ToolTipText     =   "Input employee's position in the company."
         Top             =   2550
         Width           =   3585
      End
      Begin VB.TextBox txtSignName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1380
         TabIndex        =   5
         Text            =   "Text1"
         ToolTipText     =   "Type the employee's sign name in this format: firstname MI lastname."
         Top             =   2160
         Width           =   3585
      End
      Begin VB.TextBox txtFullName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1380
         TabIndex        =   4
         Text            =   "Text1"
         ToolTipText     =   "Type the employee's full name, lastname first. "
         Top             =   1770
         Width           =   3585
      End
      Begin VB.TextBox txtMiddleInt 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
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
         Text            =   "Text1"
         ToolTipText     =   "Type the employee's middle initial"
         Top             =   1380
         Width           =   465
      End
      Begin VB.TextBox txtFirstName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
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
         Text            =   "Text1"
         ToolTipText     =   "Type the employee's first name."
         Top             =   990
         Width           =   3585
      End
      Begin VB.TextBox txtEmpNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
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
         Text            =   "Text1"
         ToolTipText     =   "Type the employee number (e.g. 0001)"
         Top             =   210
         Width           =   1755
      End
      Begin VB.TextBox txtLastName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
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
         Text            =   "Text1"
         ToolTipText     =   "Type the employee's last name."
         Top             =   600
         Width           =   3585
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   4
         Left            =   5010
         TabIndex        =   23
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   1
         Left            =   5010
         TabIndex        =   22
         Top             =   1020
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   0
         Left            =   5010
         TabIndex        =   18
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   2
         Left            =   3210
         TabIndex        =   17
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   16
         Top             =   2580
         Width           =   1275
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sign Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   15
         Top             =   2190
         Width           =   1275
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   14
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Initial"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   13
         Top             =   1410
         Width           =   1275
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   12
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Emp. No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   11
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   8
         Top             =   630
         Width           =   1275
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3945
      Left            =   60
      TabIndex        =   19
      Top             =   0
      Width           =   2595
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
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   150
         Width           =   2445
      End
      Begin MSComctlLib.ListView lstSalesMan 
         Height          =   3375
         Left            =   60
         TabIndex        =   21
         Top             =   510
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   5953
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
         MouseIcon       =   "SalesMan.frx":2D71
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FULL NAME"
            Object.Width           =   5292
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
      Height          =   855
      Left            =   6990
      ScaleHeight     =   855
      ScaleWidth      =   1470
      TabIndex        =   33
      Top             =   3090
      Width           =   1470
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
         MouseIcon       =   "SalesMan.frx":2ED3
         MousePointer    =   99  'Custom
         Picture         =   "SalesMan.frx":3025
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   735
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
         MouseIcon       =   "SalesMan.frx":3363
         MousePointer    =   99  'Custom
         Picture         =   "SalesMan.frx":34B5
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Save this Record"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.TextBox txtCode 
      Height          =   495
      Left            =   1200
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   3360
      Width           =   1215
   End
   Begin Crystal.CrystalReport RPTSALESMAN 
      Left            =   600
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   60
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   390
      TabIndex        =   9
      Top             =   210
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmPMISMaster_SalesMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSSALESMAN                                         As ADODB.Recordset
Dim AddorEdit                                          As String

Sub initMemvars()
    txtEmpNo.Text = ""
    txtLastName.Text = ""
    txtFirstName.Text = ""
    txtMiddleInt.Text = ""
    txtFullName.Text = ""
    txtSignName.Text = ""
    txtPositions.Text = ""
End Sub

Sub StoreMemVars()
    If Not RSSALESMAN.EOF And Not RSSALESMAN.BOF Then
        labid.Caption = Null2String(RSSALESMAN!empno)
        txtEmpNo.Text = Null2String(RSSALESMAN!empno)
        txtLastName.Text = Null2String(RSSALESMAN!lastname)
        txtFirstName.Text = Null2String(RSSALESMAN!FIRSTNAME)
        txtMiddleInt.Text = Null2String(RSSALESMAN!middleint)
        'txtMiddleInt.Text = Null2String(rsSalesMan!middlename)
        txtFullName.Text = Null2String(RSSALESMAN!FullName)
        txtSignName.Text = Null2String(RSSALESMAN!signname)
        txtPositions.Text = Null2String(RSSALESMAN!Positions)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set RSSALESMAN = New ADODB.Recordset
    RSSALESMAN.Open "select * from PMIS_vw_SalesMan", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub FillGrid()
    Dim rsSMan                                         As ADODB.Recordset
    lstSalesMan.Enabled = False
    lstSalesMan.Sorted = False: lstSalesMan.ListItems.Clear
    Set rsSMan = New ADODB.Recordset


    Set rsSMan = gconDMIS.Execute("select FullName,EmpNo from PMIS_vw_SalesMan  order by FullName asc")
    If Not (rsSMan.EOF And rsSMan.BOF) Then
        lstSalesMan.Enabled = True
        Listview_Loadval Me.lstSalesMan.ListItems, rsSMan
        lstSalesMan.Refresh

    Else
        lstSalesMan.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsSMan                                         As ADODB.Recordset
    lstSalesMan.Enabled = False
    lstSalesMan.Sorted = False: lstSalesMan.ListItems.Clear
    Set rsSMan = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsSMan = gconDMIS.Execute("select FullName,EmpNo from PMIS_vw_SalesMan  where FullName like'" & XXX & "%'")
    If Not (rsSMan.EOF And rsSMan.BOF) Then
        lstSalesMan.Enabled = True
        Listview_Loadval Me.lstSalesMan.ListItems, rsSMan
        lstSalesMan.Refresh
    Else
        lstSalesMan.Enabled = False
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "SALESMAN MASTER FILE") = False Then Exit Sub
    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    lstSalesMan.Enabled = False
    textSearch.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstSalesMan.Enabled = True
    textSearch.Enabled = True
    fraDetails.Enabled = True
    StoreMemVars

End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "SALESMAN MASTER FILE") = False Then Exit Sub
    On Error GoTo ErrorCode
    If Not RSSALESMAN.BOF Or Not RSSALESMAN.EOF Then
        If ShowConfirmDelete = True Then
            If Null2String(RSSALESMAN!ENTFROM) = "HRMS" Then
                MsgBox "Cannot delete this record! Please Contact Your HRD.", vbInformation
            Else
                SQL_STATEMENT = "delete from PMIS_SalesMan where EmpNo = '" & labid.Caption & "'"
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "X", "SALESMAN MASTER FILE", SQL_STATEMENT, labid, "Salesman", txtEmpNo & " - " & txtLastName, "", ""
                ShowDeletedMsg
                rsRefresh
                StoreMemVars
                FillGrid
            End If
        End If
    Else
        ShowNothingToDeleteMsg
    End If

    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "SALESMAN MASTER FILE") = False Then Exit Sub
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next

    textSearch.SetFocus

End Sub

Private Sub cmdNext_Click()
    RSSALESMAN.MoveNext
    If RSSALESMAN.EOF Then
        RSSALESMAN.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    RSSALESMAN.MovePrevious
    If RSSALESMAN.BOF Then
        RSSALESMAN.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub
'kevin 06-05-2014'
Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "SALESMAN MASTER FILE") = False Then Exit Sub
 On Error GoTo ErrorCode:

    Screen.MousePointer = 11
    RPTSALESMAN.Reset

    RPTSALESMAN.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    RPTSALESMAN.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport RPTSALESMAN, PMIS_REPORT_PATH & "Salesman.rpt", "", DMIS_REPORT_Connection, 1

    Screen.MousePointer = 0
    LogAudit "S", "SALESMAN MASTER FILE", txtCode
    Exit Sub
ErrorCode:
    ShowVBError
    Screen.MousePointer = 0



End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode
    Dim RSFINDDUP                                      As ADODB.Recordset

    Dim VTXTEmpNo, VTXTLASTNAME, VTXTFIRSTNAME         As String
    Dim VTXTMiddleInt, VTXTFullName, VTXTSignName      As String
    Dim VTXTPositions                                  As String

    If IsNull(txtEmpNo.Text) = True Then
        MsgSpeechBox "Employee Number must not be empty"
        On Error Resume Next

        txtEmpNo.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set RSFINDDUP = New ADODB.Recordset
            RSFINDDUP.Open "select empno from PMIS_vw_SalesMan where empno = '" & txtEmpNo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not RSFINDDUP.EOF And Not RSFINDDUP.BOF Then
                MsgSpeechBox "Employee Number already exist!"
                On Error Resume Next
                txtEmpNo.SetFocus
                Exit Sub
            End If
        ElseIf AddorEdit = "EDIT" And txtEmpNo <> Null2String(RSSALESMAN!empno) Then
            Set RSFINDDUP = New ADODB.Recordset
            RSFINDDUP.Open "select empno from PMIS_vw_SalesMan where empno = '" & txtEmpNo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not RSFINDDUP.EOF And Not RSFINDDUP.BOF Then
                MsgSpeechBox "Employee Number already exist!"
                On Error Resume Next
                txtEmpNo.SetFocus
                Exit Sub
            End If


        End If


    End If
    If txtLastName.Text = "" Then
        ShowIsRequiredMsg "Last Name"
        Exit Sub
    End If

    VTXTEmpNo = N2Str2Null(txtEmpNo.Text)
    VTXTLASTNAME = N2Str2Null(txtLastName.Text)
    VTXTFIRSTNAME = N2Str2Null(txtFirstName.Text)
    VTXTMiddleInt = N2Str2Null(txtMiddleInt.Text)
    VTXTFullName = N2Str2Null(txtFullName.Text)
    VTXTSignName = N2Str2Null(txtSignName.Text)
    VTXTPositions = N2Str2Null(txtPositions.Text)

    If AddorEdit = "ADD" Then

        SQL_STATEMENT = "Insert into PMIS_SalesMan" & _
                      " (empno,lastname,firstname,middleint,fullname,signname,Positions,lastUpdate,UserCode)" & _
                      " values (" & VTXTEmpNo & ", " & VTXTLASTNAME & _
                        ", " & VTXTFIRSTNAME & ", " & VTXTMiddleInt & _
                        ", " & VTXTFullName & ", " & VTXTSignName & _
                        ", " & VTXTPositions & _
                        ", " & "'" & LOGDATE & "'" & ", " & _
                      " " & "" & N2Str2Null(LOGCODE) & "" & ")"
        gconDMIS.Execute SQL_STATEMENT

        NEW_LogAudit "A", "SALESMAN MASTER FILE", SQL_STATEMENT, labid, "Salesman", txtEmpNo & " - " & txtLastName, "", ""

        ShowSuccessFullyAdded
    Else
        If Null2String(RSSALESMAN!ENTFROM) = "HRMS" Then

            SQL_STATEMENT = "update HRMS_EmpInfo set" & _
                          " empno = " & VTXTEmpNo & "," & _
                          " Lastname = " & VTXTLASTNAME & "," & _
                          " Firstname = " & VTXTFIRSTNAME & "," & _
                          " Middlename = " & VTXTMiddleInt & "," & _
                          " [position] = " & VTXTPositions & "," & _
                          " LastUpdate = " & "'" & LOGDATE & "'" & "," & _
                          " UserCode = " & "" & N2Str2Null(LOGCODE) & "" & _
                          " where EmpNo = '" & labid.Caption & "'"
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "E", "SALESMAN MASTER FILE", SQL_STATEMENT, labid, "Salesman", txtEmpNo & " - " & txtLastName, "", ""

        Else
            SQL_STATEMENT = "update PMIS_SalesMan set" & _
                          " empno = " & VTXTEmpNo & "," & _
                          " lastname = " & VTXTLASTNAME & "," & _
                          " firstname = " & VTXTFIRSTNAME & "," & _
                          " middleint = " & VTXTMiddleInt & "," & _
                          " fullname = " & VTXTFullName & "," & _
                          " signname = " & VTXTSignName & "," & _
                          " Positions = " & VTXTPositions & "," & _
                          " LastUpdate = " & "'" & LOGDATE & "'" & "," & _
                          " UserCode = " & "" & N2Str2Null(LOGCODE) & "" & _
                          " where EmpNo = '" & labid.Caption & "'"
            gconDMIS.Execute SQL_STATEMENT




            NEW_LogAudit "E", "SALESMAN MASTER FILE", SQL_STATEMENT, labid, "Salesman", txtEmpNo & " - " & txtLastName, "", ""
        End If


        ShowSuccessFullyUpdated
    End If
    rsRefresh
    FillGrid
    On Error Resume Next
    If AddorEdit = "EDIT" Then
        RSSALESMAN.Find "empno = '" & txtEmpNo & "'"
    End If

    cmdCancel.Value = True
    Exit Sub

ErrorCode:
    'MsgBox err.Description
    ShowVBError
    cmdCancel.Value = True
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    Frame1.Enabled = False
    textSearch.Text = ""
    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISMaster_SalesMan = Nothing
    UnloadForm Me
End Sub

'Update by: NVB -----------------------------------
'Commented by NVB
'Descprition: Show to the listview  the new salesman, and also show the info of salesman
'Private Sub lstSalesMan_GotFocus()
'    rsSalesMan.Bookmark = rsFind(rsSalesMan.Clone, "FULLNAME", lstSalesMan.SelectedItem.SubItems(0))).Bookmark
'    StoreMemvars
'End Sub

Private Sub lstSalesMan_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'rsSalesMan.Bookmark = rsFind(rsSalesMan.Clone, "EMPNO", lstSalesMan.SelectedItem.SubItems(0))).Bookmark
    RSSALESMAN.Bookmark = rsFind(RSSALESMAN.Clone, "FULLNAME", Item).Bookmark
    StoreMemVars

End Sub

' ------------------------------------------------------

Private Sub lstSalesMan_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstSalesMan
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

Private Sub lstSalesMan_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstSalesMan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub Text1_Change()

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
        If lstSalesMan.ListItems.Count > 0 And lstSalesMan.Enabled = True Then: lstSalesMan.SetFocus
    End If
End Sub

Private Sub txtFirstName_Change()
    txtLastName_Change
End Sub

Private Sub txtLastName_Change()
    txtFullName = UCase(txtLastName + ", " + txtFirstName + " " + Left(txtMiddleInt, 1) + ".")
    txtSignName = UCase(txtLastName + ", " + txtFirstName + " " + Left(txtMiddleInt, 1) + ".")
End Sub

Private Sub txtMiddleInt_Change()
    txtLastName_Change
End Sub


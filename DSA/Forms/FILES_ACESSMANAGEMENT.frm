VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmFiles_AcessManagement 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "USERS ACCESS"
   ClientHeight    =   8520
   ClientLeft      =   315
   ClientTop       =   570
   ClientWidth     =   14325
   ClipControls    =   0   'False
   ForeColor       =   &H00F5F5F5&
   Icon            =   "FILES_ACESSMANAGEMENT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   14325
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   3360
      ScaleHeight     =   915
      ScaleWidth      =   8085
      TabIndex        =   19
      Top             =   7530
      Visible         =   0   'False
      Width           =   8115
      Begin wizProgBar.Prg Prg1 
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   873
         Picture         =   "FILES_ACESSMANAGEMENT.frx":0E42
         ForeColor       =   0
         BarPicture      =   "FILES_ACESSMANAGEMENT.frx":0E5E
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
      Begin VB.Label labPercentage 
         Alignment       =   1  'Right Justify
         Caption         =   "0 %"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6510
         TabIndex        =   22
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label labModuleName 
         Caption         =   "Module Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   90
         Width           =   5115
      End
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   14325
      TabIndex        =   0
      Top             =   0
      Width           =   14325
      Begin VB.ComboBox cboUsers 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   2535
      End
      Begin VB.ComboBox cboMainModule 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "FILES_ACESSMANAGEMENT.frx":0E7A
         Left            =   2730
         List            =   "FILES_ACESSMANAGEMENT.frx":0E7C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   2565
      End
      Begin VB.ComboBox cboModuleType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "FILES_ACESSMANAGEMENT.frx":0E7E
         Left            =   5370
         List            =   "FILES_ACESSMANAGEMENT.frx":0E80
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   300
         Width           =   3525
      End
      Begin VB.CommandButton Command1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8910
         Picture         =   "FILES_ACESSMANAGEMENT.frx":0E82
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   1
         Top             =   30
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "System "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2730
         TabIndex        =   2
         ToolTipText     =   "System That User Can Access"
         Top             =   60
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Module Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5340
         TabIndex        =   3
         Top             =   60
         Width           =   1065
      End
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
      Height          =   405
      Left            =   2670
      MouseIcon       =   "FILES_ACESSMANAGEMENT.frx":1301
      MousePointer    =   99  'Custom
      Picture         =   "FILES_ACESSMANAGEMENT.frx":1453
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   750
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      MaxLength       =   100
      TabIndex        =   13
      Top             =   1200
      Width           =   3225
   End
   Begin MSComctlLib.ListView lvwModules 
      Height          =   6870
      Left            =   30
      TabIndex        =   14
      Top             =   1590
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   12118
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
      MouseIcon       =   "FILES_ACESSMANAGEMENT.frx":1735
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "MODULES"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "id"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   11490
      ScaleHeight     =   885
      ScaleWidth      =   4680
      TabIndex        =   15
      Top             =   7530
      Width           =   4680
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
         Left            =   2070
         MouseIcon       =   "FILES_ACESSMANAGEMENT.frx":1897
         MousePointer    =   99  'Custom
         Picture         =   "FILES_ACESSMANAGEMENT.frx":19E9
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   90
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1380
         MouseIcon       =   "FILES_ACESSMANAGEMENT.frx":1D4F
         MousePointer    =   99  'Custom
         Picture         =   "FILES_ACESSMANAGEMENT.frx":1EA1
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   90
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
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
         Left            =   690
         MouseIcon       =   "FILES_ACESSMANAGEMENT.frx":2340
         MousePointer    =   99  'Custom
         Picture         =   "FILES_ACESSMANAGEMENT.frx":2492
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   90
         Width           =   705
      End
   End
   Begin VB.PictureBox picGrid 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   3330
      ScaleHeight     =   6735
      ScaleWidth      =   13695
      TabIndex        =   9
      Top             =   810
      Width           =   13695
      Begin Crystal.CrystalReport rtpPrint 
         Left            =   3000
         Top             =   5730
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CheckBox Check1 
         Caption         =   "CHECK ALL "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6840
         TabIndex        =   11
         Top             =   60
         Width           =   1785
      End
      Begin FlexCell.Grid Grid 
         Height          =   6315
         Left            =   30
         TabIndex        =   24
         Top             =   390
         Width           =   10905
         _ExtentX        =   19235
         _ExtentY        =   11139
         Cols            =   6
         DefaultFontSize =   8.25
         DisplayRowIndex =   -1  'True
         GridColor       =   12632256
         Rows            =   1
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "** PLEASE PRESS SAVE AFTER EDITING CELL(s)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Left            =   30
         TabIndex        =   23
         Top             =   6060
         Width           =   4005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User can access the ff. modules:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   90
         TabIndex        =   10
         Top             =   30
         Width           =   2670
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   930
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuExportSettings 
         Caption         =   "Export Setting"
      End
      Begin VB.Menu mnuImportSettings 
         Caption         =   "Import Settings"
      End
   End
End
Attribute VB_Name = "frmFiles_AcessManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NameStr                             As String
Dim mUserID                             As Long
Dim SQL                                 As String
Dim i                                   As Integer
Dim TEMPRS                              As ADODB.Recordset
Public Property Let UserID(sUserID As Long)
    mUserID = sUserID
End Property

Public Property Let Username(sNamestr As String)
    NameStr = sNamestr
End Property

Private Sub cboMainModule_CLICK()
    Command1.Enabled = True
    cmdPrint.Enabled = False
End Sub

Private Sub FillData()
    If cboModuleType.ListIndex = -1 Then: Exit Sub
    If cboMainModule.ListIndex = -1 Then Exit Sub
    Dim RS                              As ADODB.Recordset
    Dim NotInSQL                        As String

    NotInSQL = "SELECT   DESCRIPTIONS,  MODULEID  FROM ALL_RAMS_MODULES " & _
             " WHERE    MODULEID NOT IN(SELECT MODULEID FROM ALL_vW_RAMS_USERACESS  WHERE userid = " & mUserID & ") " & _
             "  and  MAINMODULENAME = '" & cboMainModule.Text & "' and MODULE_TYPE='" & UCase(cboModuleType.Text) & "' Order by DESCRIPTIONS"

    Set RS = gconDMIS.Execute(NotInSQL)
    'Stop
    If Not (RS.BOF And RS.EOF) Then
        lvwModules.Enabled = True
        Listview_Loadval Me.lvwModules.ListItems, RS
    Else
        lvwModules.Enabled = False
        Me.lvwModules.ListItems.Clear
    End If

    Set RS = Nothing

    Screen.MousePointer = 11
    i = 0
    Grid.Visible = False
    '    Stop
    Select Case UCase(cboModuleType.Text)
        Case "SYSTEM"
            ShowACCESS_System cboMainModule.Text, mUserID, cboModuleType.Text
        Case "DATA ENTRY"
            ShowACCESS_DataEntry cboMainModule.Text, mUserID, cboModuleType.Text
        Case "SEARCH"
            ShowACCESS_SEARCH cboMainModule.Text, mUserID, cboModuleType.Text
        Case "INQUIRY"
            ShowACCESS_INQUIRY cboMainModule.Text, mUserID, cboModuleType.Text
        Case "REPORTS"
            ShowACCESS_Reports cboMainModule.Text, mUserID, cboModuleType.Text
        Case "PROCESSING"
            ShowACCESS_Processing cboMainModule.Text, mUserID, cboModuleType.Text
        Case "TRANSACTION"
            ShowACCESS_TRANSACTION cboMainModule.Text, mUserID, cboModuleType.Text
    End Select
    Grid.Visible = True
    Grid.Refresh
    Set TEMPRS = Nothing
    SQL = vbNullString
    '    If Grid.Rows > 1 Then
    '        cboModuleType.Enabled = False
    '        cboMainModule.Enabled = False
    '
    '    Else
    '        cboModuleType.Enabled = True
    '        cboMainModule.Enabled = True
    '
    '    End If
    '
    Screen.MousePointer = 0
End Sub

Sub AddHyperLink(colINDEX As Integer)
    With Grid
        .Column(1).Locked = True
        .Column(colINDEX).Locked = True
        .Column(colINDEX).CellType = cellTextBox
        .Column(colINDEX).Width = 50
        .Cell(0, colINDEX).Text = "OPTIONS"
    End With
End Sub

Sub ShowACCESS_System(XMainModuleName, XUserID, XmoduleType)
    SQL = "SELECT ARU.MODULEID, ARM.DESCRIPTIONS,  ARU.Acess_System  " & vbCrLf & _
        " FROM ALL_RAMS_USERSACESS ARU INNER JOIN  ALL_RAMS_MODULES  ARM ON ARU.MODULEID = ARM.MODULEID " & vbCrLf & _
        " WHERE (ARM.MAINMODULENAME = '" & XMainModuleName & "' AND ARU.USERID = " & XUserID & " AND ARM.MODULE_TYPE = '" & XmoduleType & "') "
    With Grid
        .Cols = 4
        .FixedCols = 1
        .Column(1).Width = 281
        .Column(2).Width = 0
        .Cell(0, 1).Text = "MODULE NAME"
        .Cell(0, 2).Text = "ACCESS"

        .Column(2).CellType = cellCheckBox
        Call AddHyperLink(3)
    End With

    Set TEMPRS = gconDMIS.Execute(SQL)
    Grid.Rows = 1
    While Not TEMPRS.EOF
        With Grid
            i = i + 1
            .AddItem TEMPRS!DESCRIPTIONS, False
            .Cell(i, 0).Text = Null2String(TEMPRS!ModuleID)
            .Cell(i, 2).Text = Null2String(TEMPRS!Acess_System)
            .Cell(i, 3).Font.Bold = True
            .Cell(i, 3).ForeColor = vbBlue
            .Cell(i, 3).Text = "DELETE"
        End With
        TEMPRS.MoveNext
    Wend
    Grid.Refresh

End Sub


Sub ShowACCESS_DataEntry(XMainModuleID, XUserID, XmoduleType)

    SQL = "SELECT ARU.MODULEID, ARM.DESCRIPTIONS, ARU.Acess_Add, ARU.Acess_Edit, ARU.Acess_Delete, ARU.Acess_View, ARU.Acess_Print, ARU.Acess_Process  " & vbCrLf & _
        " FROM ALL_RAMS_USERSACESS ARU INNER JOIN  ALL_RAMS_MODULES  ARM ON ARU.MODULEID = ARM.MODULEID " & vbCrLf & _
        " WHERE (ARM.MAINMODULENAME = '" & XMainModuleID & "'  AND ARU.USERID = " & XUserID & " AND ARM.MODULE_TYPE='" & XmoduleType & "' ) "
    With Grid
        .Cols = 8
        .Column(1).Width = 190
        .Column(2).Width = 50
        .Column(3).Width = 50
        .Column(4).Width = 0
        .Column(5).Width = 50
        .Column(6).Width = 50
        .Column(2).CellType = cellCheckBox
        .Column(3).CellType = cellCheckBox
        .Column(4).CellType = cellCheckBox
        .Column(5).CellType = cellCheckBox
        .Column(6).CellType = cellCheckBox
        Call AddHyperLink(7)
        .Cell(0, 1).Text = "MODULE NAME"
        .Cell(0, 2).Text = "ADD"
        .Cell(0, 3).Text = "EDIT"
        .Cell(0, 4).Text = "VIEW"
        .Cell(0, 5).Text = "DELETE"
        .Cell(0, 6).Text = "PRINT"


    End With
    Set TEMPRS = gconDMIS.Execute(SQL)

    Grid.Rows = 1
    While Not TEMPRS.EOF
        With Grid
            i = i + 1
            .AddItem TEMPRS!DESCRIPTIONS, False
            .Cell(i, 7).ForeColor = vbBlue
            .Cell(i, 7).Text = "DELETE"
            .Cell(i, 7).Font.Bold = True
            .Cell(i, 0).Text = (TEMPRS!ModuleID)
            .Cell(i, 2).Text = Null2String(TEMPRS!Acess_Add)
            .Cell(i, 3).Text = Null2String(TEMPRS!Acess_Edit)
            .Cell(i, 4).Text = Null2String(TEMPRS!Acess_View)
            .Cell(i, 5).Text = Null2String(TEMPRS!Acess_Delete)
            .Cell(i, 6).Text = Null2String(TEMPRS!Acess_Print)
        End With
        TEMPRS.MoveNext
    Wend

    Grid.Refresh

End Sub
Sub ShowACCESS_SEARCH(XMainModuleID, XUserID, XmoduleType)
    SQL = "SELECT ARU.MODULEID, ARM.DESCRIPTIONS, ARU.Acess_System " & vbCrLf & _
        " FROM ALL_RAMS_USERSACESS ARU INNER JOIN  ALL_RAMS_MODULES  ARM ON ARU.MODULEID = ARM.MODULEID " & vbCrLf & _
        " WHERE (ARM.MAINMODULENAME = '" & XMainModuleID & "'  AND ARU.USERID = " & XUserID & " AND ARM.MODULE_TYPE='" & XmoduleType & "' ) "
    With Grid
        .Cols = 4
        .Column(1).Width = 281
        .Column(2).Width = 75
        .Column(2).CellType = cellCheckBox
        AddHyperLink (3)
        .Cell(0, 1).Text = "MODULE NAME"
        .Cell(0, 2).Text = "ACCESS"

    End With
    Set TEMPRS = gconDMIS.Execute(SQL)

    Grid.Rows = 1
    While Not TEMPRS.EOF
        With Grid
            i = i + 1
            .AddItem TEMPRS!DESCRIPTIONS, False
            .Cell(i, 0).Text = Null2String(TEMPRS!ModuleID)
            .Cell(i, 2).Text = Null2String(TEMPRS!Acess_System)
            .Cell(i, 3).ForeColor = vbBlue
            .Cell(i, 3).Text = "DELETE"
            .Cell(i, 3).Font.Bold = True
        End With

        TEMPRS.MoveNext
    Wend

    Grid.Refresh
End Sub


Sub ShowACCESS_INQUIRY(XMainModuleID, XUserID, XmoduleType)
    SQL = "SELECT ARU.MODULEID, ARM.DESCRIPTIONS, ARU.Acess_System " & vbCrLf & _
        " FROM ALL_RAMS_USERSACESS ARU INNER JOIN  ALL_RAMS_MODULES  ARM ON ARU.MODULEID = ARM.MODULEID " & vbCrLf & _
        " WHERE (ARM.MAINMODULENAME = '" & XMainModuleID & "'  AND ARU.USERID = " & XUserID & " AND ARM.MODULE_TYPE='" & XmoduleType & "' ) "
    With Grid
        .Cols = 4
        .Column(1).Width = 281
        .Column(2).Width = 0
        .Column(2).CellType = cellCheckBox
        .Cell(0, 1).Text = "MODULE NAME"
        .Cell(0, 2).Text = "ACCESS"
        AddHyperLink (3)
    End With
    Set TEMPRS = gconDMIS.Execute(SQL)

    Grid.Rows = 1
    While Not TEMPRS.EOF
        With Grid
            i = i + 1
            .AddItem TEMPRS!DESCRIPTIONS, False
            .Cell(i, 0).Text = Null2String(TEMPRS!ModuleID)
            .Cell(i, 2).Text = Null2String(TEMPRS!Acess_System)
            .Cell(i, 3).ForeColor = vbBlue
            .Cell(i, 3).Text = "DELETE"
            .Cell(i, 3).Font.Bold = True
        End With

        TEMPRS.MoveNext
    Wend

    Grid.Refresh
End Sub

Sub ShowACCESS_Reports(XMainModuleID, XUserID, XmoduleType)
    SQL = "SELECT ARU.MODULEID, ARM.DESCRIPTIONS, ARU.Acess_View, ARU.Acess_Print  " & vbCrLf & _
        " FROM ALL_RAMS_USERSACESS ARU INNER JOIN  ALL_RAMS_MODULES  ARM ON ARU.MODULEID = ARM.MODULEID " & vbCrLf & _
        " WHERE (ARM.MAINMODULENAME = '" & XMainModuleID & "'  AND ARU.USERID = " & XUserID & " AND ARM.MODULE_TYPE = '" & XmoduleType & "') "
    With Grid
        .Cols = 5
        .Column(1).Width = 281
        .Column(2).Width = 0
        .Column(3).Width = 0
        .Column(2).CellType = cellCheckBox
        .Column(3).CellType = cellCheckBox
        .Cell(0, 1).Text = "MODULE NAME"
        .Cell(0, 2).Text = "EXPORT"
        .Cell(0, 3).Text = "PRINT"
        AddHyperLink (4)
    End With
    Set TEMPRS = gconDMIS.Execute(SQL)

    Grid.Rows = 1
    While Not TEMPRS.EOF
        With Grid
            i = i + 1
            .AddItem TEMPRS!DESCRIPTIONS, False
            .Cell(i, 0).Text = Null2String(TEMPRS!ModuleID)
            .Cell(i, 2).Text = Null2String(TEMPRS!Acess_View)
            .Cell(i, 3).Text = Null2String(TEMPRS!Acess_Print)
            .Cell(i, 4).ForeColor = vbBlue
            .Cell(i, 4).Text = "DELETE"
            .Cell(i, 4).Font.Bold = True
        End With
        TEMPRS.MoveNext
    Wend

    Grid.Refresh
End Sub

Sub ShowACCESS_Processing(XMainModuleID, XUserID, XmoduleType)
    SQL = "SELECT ARU.MODULEID, ARM.DESCRIPTIONS, ARU.Acess_Process" & vbCrLf & _
        " FROM ALL_RAMS_USERSACESS ARU INNER JOIN  ALL_RAMS_MODULES  ARM ON ARU.MODULEID = ARM.MODULEID " & vbCrLf & _
        " WHERE (ARM.MAINMODULENAME = '" & XMainModuleID & "'  AND ARU.USERID = " & XUserID & " AND ARM.MODULE_TYPE = '" & XmoduleType & "') "

    With Grid
        .Cols = 4
        .Column(1).Width = 281
        .Column(2).Width = 0
        .Column(2).CellType = cellCheckBox
        AddHyperLink (3)
        .Cell(0, 1).Text = "MODULE NAME"
        .Cell(0, 2).Text = "IMPORT/PROCESS"


    End With
    Set TEMPRS = gconDMIS.Execute(SQL)

    Grid.Rows = 1
    While Not TEMPRS.EOF
        With Grid
            i = i + 1
            .AddItem TEMPRS!DESCRIPTIONS, False
            .Cell(i, 0).Text = Null2String(TEMPRS!ModuleID)
            .Cell(i, 2).Text = Null2String(TEMPRS!Acess_Process)
            .Cell(i, 3).Text = "DELETE"
            .Cell(i, 3).ForeColor = vbBlue
            .Cell(i, 3).Font.Bold = True
        End With
        TEMPRS.MoveNext
    Wend

    Grid.Refresh
End Sub
Sub ShowACCESS_TRANSACTION(XMainModuleID, XUserID, XmoduleType)
    SQL = "SELECT  ARU.MODULEID, ARM.DESCRIPTIONS, ARU.Acess_Add, ARU.Acess_Edit, ARU.Acess_Delete, ARU.Acess_CancelEntry, ARU.Acess_Print, ARU.Acess_Post, ARU.Acess_UnPost " & vbCrLf & _
          ", ARU.Acess_System , ARU.Acess_Reprint FROM ALL_RAMS_USERSACESS ARU INNER JOIN  ALL_RAMS_MODULES  ARM ON ARU.MODULEID = ARM.MODULEID " & vbCrLf & _
        " WHERE (ARM.MAINMODULENAME = '" & XMainModuleID & "'  AND ARU.USERID = " & XUserID & " AND ARM.MODULE_TYPE = '" & XmoduleType & "') "

    With Grid
        .Cols = 12
        .Column(1).Width = 220
        .Column(2).Width = 40
        .Column(3).Width = 40
        .Column(4).Width = 40
        .Column(5).Width = 0
        .Column(6).Width = 40
        .Column(7).Width = 50
        .Column(8).Width = 50
        .Column(9).Width = 80
        .Column(10).Width = 50
        
        .Column(2).CellType = cellCheckBox
        .Column(3).CellType = cellCheckBox
        .Column(4).CellType = cellCheckBox
        .Column(5).CellType = cellCheckBox
        .Column(6).CellType = cellCheckBox
        .Column(7).CellType = cellCheckBox
        .Column(8).CellType = cellCheckBox
        .Column(9).CellType = cellCheckBox
        .Column(10).CellType = cellCheckBox
        AddHyperLink (11)

        .Cell(0, 1).Text = "MODULE NAME"

        .Cell(0, 2).Font.Size = 7
        .Cell(0, 2).Text = "ADD"

        .Cell(0, 3).Font.Size = 7
        .Cell(0, 3).Text = "EDIT"

        .Cell(0, 4).Font.Size = 7
        .Cell(0, 4).Text = "PRINT"

        .Cell(0, 5).Font.Size = 7
        .Cell(0, 5).Text = "DELETE"

        .Cell(0, 6).Font.Size = 7
        .Cell(0, 6).Text = "POST"

        .Cell(0, 7).Font.Size = 7
        .Cell(0, 7).Text = "UNPOST"

        .Cell(0, 8).Font.Size = 7
        .Cell(0, 8).Text = "CANCEL"

        .Cell(0, 9).Font.Size = 7
        .Cell(0, 9).Text = "EDIT TRANDATE"
        
        .Cell(0, 10).Font.Size = 7
        .Cell(0, 10).Text = "RE-PRINT"
        
        .Cell(0, 11).Font.Size = 8
        
        
    End With
    Set TEMPRS = gconDMIS.Execute(SQL)

    Grid.Rows = 1
    While Not TEMPRS.EOF
        With Grid
            i = i + 1
            .AddItem TEMPRS!DESCRIPTIONS, False
            .Cell(i, 0).Text = Null2String(TEMPRS!ModuleID)
            .Cell(i, 2).Text = Null2String(TEMPRS!Acess_Add)
            .Cell(i, 3).Text = Null2String(TEMPRS!Acess_Edit)
            .Cell(i, 4).Text = Null2String(TEMPRS!Acess_Print)
            .Cell(i, 5).Text = Null2String(TEMPRS!Acess_Delete)
            .Cell(i, 6).Text = Null2String(TEMPRS!Acess_POST)
            .Cell(i, 7).Text = Null2String(TEMPRS!Acess_UnPost)
            .Cell(i, 8).Text = Null2String(TEMPRS!Acess_CancelEntry)
            .Cell(i, 9).Text = Null2String(TEMPRS!Acess_System)
            .Cell(i, 10).Text = Null2String(TEMPRS!Acess_RePrint)
            .Cell(i, 11).ForeColor = vbBlue
            .Cell(i, 11).Text = "DELETE"
            .Cell(i, 11).Font.Bold = True
        End With
        TEMPRS.MoveNext
    Wend

    Grid.Refresh

End Sub

Private Sub Save_Modules()
    Dim TEMPSQL                         As String
    Dim lxJ                             As Long
    Dim GMax As Integer
    
    On Error GoTo ErrorCode:
    If mUserID = 0 Then
        MsgBox "NO USR", vbInformation
        Exit Sub
    End If
    
    picProgress.Visible = True: picProgress.ZOrder 0
    Prg1.Value = 0
    gconDMIS.Execute ("Delete from ALL_RAMS_USERSACESS where USERID=" & mUserID & " AND  MODULEID IN (Select MODULEID from ALL_vW_RAMS_USERACESS WHERE  MODULE_TYPE='" & UCase(cboModuleType.Text) & "' AND USERID=" & mUserID & " AND MAINMODULENAME='" & cboMainModule.Text & "')")
    GMax = Grid.Rows - 1
    Prg1.Max = GMax
    labModuleName = ""
    labPercentage = "0%"
    DoEvents
    Select Case UCase(cboModuleType.Text)
        Case "DATA ENTRY"
            For lxJ = 1 To Grid.Rows - 1
            Prg1.Value = lxJ
                labPercentage = FormatPercent(lxJ / GMax)
                labModuleName = Grid.Cell(lxJ, 1).Text
                 
                TEMPSQL = " INSERT INTO ALL_RAMS_USERSACESS" & _
                        " ( USERID, MODULEID , Acess_Add, Acess_Edit, Acess_View, Acess_Delete, Acess_Print) " & _
                        " VALUES(" & mUserID & ", " _
                        & NumericVal(Grid.Cell(lxJ, 0).Text) & ", " _
                        & NumericVal(Grid.Cell(lxJ, 2).Text) & ", " _
                        & NumericVal(Grid.Cell(lxJ, 3).Text) & ", " _
                        & NumericVal(Grid.Cell(lxJ, 4).Text) & ", " _
                        & NumericVal(Grid.Cell(lxJ, 5).Text) & ", " _
                        & NumericVal(Grid.Cell(lxJ, 6).Text) & " )"
                
                'labPercentage = Format((lxJ / GMax), 2) / 100 & "%"
                
                gconDMIS.Execute TEMPSQL

            Next
        Case "SEARCH", "SYSTEM", "INQUIRY"
            For lxJ = 1 To Grid.Rows - 1
                TEMPSQL = " INSERT INTO ALL_RAMS_USERSACESS" & _
                        " ( USERID, MODULEID , Acess_System) " & _
                        " VALUES(" & mUserID & ", " & NumericVal(Grid.Cell(lxJ, 0).Text) & ", " & NumericVal(Grid.Cell(lxJ, 2).Text) & ") "
                gconDMIS.Execute TEMPSQL
            Next
        Case "REPORTS"
            For lxJ = 1 To Grid.Rows - 1
                TEMPSQL = " INSERT INTO ALL_RAMS_USERSACESS" & _
                        " ( USERID, MODULEID , Acess_View, Acess_Print) " & _
                        " VALUES(" & mUserID & ", " & NumericVal(Grid.Cell(lxJ, 0).Text) & ", " _
                        & NumericVal(Grid.Cell(lxJ, 2).Text) & ", " _
                        & NumericVal(Grid.Cell(lxJ, 3).Text) & ") "
                gconDMIS.Execute TEMPSQL

            Next
        Case "PROCESSING"
            For lxJ = 1 To Grid.Rows - 1
                TEMPSQL = " INSERT INTO ALL_RAMS_USERSACESS" & _
                        " ( USERID, MODULEID , Acess_Process) " & _
                        " VALUES(" & mUserID & ", " & NumericVal(Grid.Cell(lxJ, 0).Text) & ", " _
                        & NumericVal(Grid.Cell(lxJ, 2).Text) & " ) "
                gconDMIS.Execute TEMPSQL

            Next
        Case "TRANSACTION"
            For lxJ = 1 To Grid.Rows - 1
                TEMPSQL = " INSERT INTO ALL_RAMS_USERSACESS" & _
                        " ( USERID, MODULEID , Acess_Add, Acess_Edit, Acess_Print ,  Acess_Delete, Acess_POST, Acess_UnPost, Acess_CancelEntry,Acess_system ,Acess_Reprint) " & _
                        " VALUES(" & mUserID & ", " _
                        & NumericVal(Grid.Cell(lxJ, 0).Text) & ", " _
                        & NumericVal(Grid.Cell(lxJ, 2).Text) & ", " _
                        & NumericVal(Grid.Cell(lxJ, 3).Text) & ", " _
                        & NumericVal(Grid.Cell(lxJ, 4).Text) & ", " _
                        & NumericVal(Grid.Cell(lxJ, 5).Text) & ", " _
                        & NumericVal(Grid.Cell(lxJ, 6).Text) & ", " _
                        & NumericVal(Grid.Cell(lxJ, 7).Text) & ", " _
                        & NumericVal(Grid.Cell(lxJ, 8).Text) & ", " _
                        & NumericVal(Grid.Cell(lxJ, 9).Text) & ", " _
                        & NumericVal(Grid.Cell(lxJ, 10).Text) & ") "
                gconDMIS.Execute TEMPSQL

            Next
    End Select

    Grid.Refresh
picProgress.Visible = False: picProgress.ZOrder 1
    Exit Sub
ErrorCode:
    'gconDMIS.RollbackTrans
    ShowVBError
End Sub



Private Sub cboModuleType_Click()
    Command1.Enabled = True
    cmdPrint.Enabled = False
End Sub

Private Sub Check1_Click()
    
        Dim j                           As Long

        For i = 2 To Grid.Cols - 2
            For j = 1 To Grid.Rows - 1
                Grid.Cell(j, i).Text = Check1.Value
            Next

        Next

        cmdSave.Enabled = True
        Grid.Refresh
    
End Sub

Private Sub cmdAdd_Click()
    If lvwModules.SelectedItem Is Nothing Then
        Exit Sub
    End If
    Dim TEMPSQL                         As String
    If Me.lvwModules.ListItems.Count > 0 Then
        Dim lngRows                     As Long

        'gconDMIS.Execute ("INSERT INTO ALL_RAMS_USERSACESS ( MODULEID, USERID) values(" & lvwModules.SelectedItem.ListSubItems(1).Text & "," & mUserID & ")")

        lngRows = Grid.Rows

        Select Case UCase(cboModuleType.Text)
            Case "SYSTEM"
                With Grid
                    .AddItem lvwModules.SelectedItem.Text, True
                    .Cell(lngRows, 0).Text = lvwModules.SelectedItem.ListSubItems(1).Text
                    .Cell(lngRows, 2).Text = 1
                    .Cell(lngRows, 3).Font.Bold = True
                    .Cell(lngRows, 3).ForeColor = vbBlue
                    .Cell(lngRows, 3).Text = "DELETE"
                End With
                TEMPSQL = " INSERT INTO ALL_RAMS_USERSACESS" & _
                        " ( USERID, MODULEID , Acess_System) " & _
                        " VALUES(" & mUserID & ", " & NumericVal(Grid.Cell(lngRows, 0).Text) & ", " & NumericVal(Grid.Cell(lngRows, 2).Text) & ") "
                gconDMIS.Execute TEMPSQL

            Case "DATA ENTRY"
                With Grid
                    .AddItem lvwModules.SelectedItem.Text, True
                    .Cell(lngRows, 0).Text = lvwModules.SelectedItem.ListSubItems(1).Text
                    .Cell(lngRows, 2).Text = 1
                    .Cell(lngRows, 3).Text = 1
                    .Cell(lngRows, 4).Text = 1
                    .Cell(lngRows, 5).Text = 1
                    .Cell(lngRows, 6).Text = 1
                    .Cell(lngRows, 7).ForeColor = vbBlue
                    .Cell(lngRows, 7).Text = "DELETE"
                    .Cell(lngRows, 7).Font.Bold = True
                End With
                TEMPSQL = " INSERT INTO ALL_RAMS_USERSACESS" & _
                        " ( USERID, MODULEID , Acess_Add, Acess_Edit, Acess_View, Acess_Delete, Acess_Print) " & _
                        " VALUES(" & mUserID & ", " _
                        & NumericVal(Grid.Cell(lngRows, 0).Text) & ", " _
                        & NumericVal(Grid.Cell(lngRows, 2).Text) & ", " _
                        & NumericVal(Grid.Cell(lngRows, 3).Text) & ", " _
                        & NumericVal(Grid.Cell(lngRows, 4).Text) & ", " _
                        & NumericVal(Grid.Cell(lngRows, 5).Text) & ", " _
                        & NumericVal(Grid.Cell(lngRows, 6).Text) & " )"
                gconDMIS.Execute TEMPSQL
            Case "SEARCH"
                With Grid
                    .AddItem lvwModules.SelectedItem.Text, True
                    .Cell(lngRows, 0).Text = lvwModules.SelectedItem.ListSubItems(1).Text
                    .Cell(lngRows, 2).Text = 1
                    .Cell(lngRows, 3).ForeColor = vbBlue
                    .Cell(lngRows, 3).Text = "DELETE"
                    .Cell(lngRows, 3).Font.Bold = True
                End With
            Case "INQUIRY"
                With Grid
                    .AddItem lvwModules.SelectedItem.Text, True
                    .Cell(lngRows, 0).Text = lvwModules.SelectedItem.ListSubItems(1).Text
                    .Cell(lngRows, 2).Text = 1
                    .Cell(lngRows, 3).ForeColor = vbBlue
                    .Cell(lngRows, 3).Text = "DELETE"
                    .Cell(lngRows, 3).Font.Bold = True
                End With

                TEMPSQL = " INSERT INTO ALL_RAMS_USERSACESS" & _
                        " ( USERID, MODULEID , Acess_System) " & _
                        " VALUES(" & mUserID & ", " & NumericVal(Grid.Cell(lngRows, 0).Text) & ", " & NumericVal(Grid.Cell(lngRows, 2).Text) & ") "
                gconDMIS.Execute TEMPSQL

            Case "REPORTS"
                With Grid
                    .AddItem lvwModules.SelectedItem.Text, True
                    .Cell(lngRows, 0).Text = lvwModules.SelectedItem.ListSubItems(1).Text
                    .Cell(lngRows, 2).Text = 1
                    .Cell(lngRows, 3).Text = 1
                    .Cell(lngRows, 4).ForeColor = vbBlue
                    .Cell(lngRows, 4).Text = "DELETE"
                    .Cell(lngRows, 4).Font.Bold = True
                End With
                TEMPSQL = " INSERT INTO ALL_RAMS_USERSACESS" & _
                        " ( USERID, MODULEID , Acess_View, Acess_Print) " & _
                        " VALUES(" & mUserID & ", " & NumericVal(Grid.Cell(lngRows, 0).Text) & ", " _
                        & NumericVal(Grid.Cell(lngRows, 2).Text) & ", " _
                        & NumericVal(Grid.Cell(lngRows, 3).Text) & ") "
                gconDMIS.Execute TEMPSQL


            Case "PROCESSING"
                With Grid
                    .AddItem lvwModules.SelectedItem.Text, True
                    .Cell(lngRows, 0).Text = lvwModules.SelectedItem.ListSubItems(1).Text
                    .Cell(lngRows, 2).Text = 1
                    .Cell(lngRows, 3).ForeColor = vbBlue
                    .Cell(lngRows, 3).Text = "DELETE"
                    .Cell(lngRows, 3).Font.Bold = True
                End With
                TEMPSQL = " INSERT INTO ALL_RAMS_USERSACESS" & _
                        " ( USERID, MODULEID , Acess_Process) " & _
                        " VALUES(" & mUserID & ", " & NumericVal(Grid.Cell(lngRows, 0).Text) & ", " _
                        & NumericVal(Grid.Cell(lngRows, 2).Text) & " ) "
                gconDMIS.Execute TEMPSQL

            Case "TRANSACTION"
                With Grid
                    '.Rows = i + 1
                    'Call .InsertRow(lngRows, 1)
                    .AddItem lvwModules.SelectedItem.Text, False
                    .Cell(lngRows, 0).Text = lvwModules.SelectedItem.ListSubItems(1).Text
                    '.Cell(lngRows, 1).Text = lvwModules.SelectedItem.Text
                    .Cell(lngRows, 2).Text = 1
                    .Cell(lngRows, 3).Text = 1
                    .Cell(lngRows, 4).Text = 1
                    .Cell(lngRows, 5).Text = 1
                    
                    .Cell(lngRows, 6).Text = 1
                    .Cell(lngRows, 7).Text = 1
                    .Cell(lngRows, 8).Text = 1
                    .Cell(lngRows, 9).Text = 1
                    .Cell(lngRows, 10).Text = 1
                    
                    .Cell(lngRows, 11).ForeColor = vbBlue
                    .Cell(lngRows, 11).Text = "DELETE"
                    .Cell(lngRows, 11).Font.Bold = True
                    .Refresh

                    TEMPSQL = " INSERT INTO ALL_RAMS_USERSACESS" & _
                            " ( USERID, MODULEID , Acess_Add, Acess_Edit, Acess_Print ,Acess_Delete, Acess_POST, Acess_UnPost, Acess_CancelEntry ,Acess_Reprint,Acess_Export,Acess_Detail) " & _
                            " VALUES(" & mUserID & ", " _
                            & NumericVal(Grid.Cell(lngRows, 0).Text) & ", " _
                            & NumericVal(Grid.Cell(lngRows, 2).Text) & ", " _
                            & NumericVal(Grid.Cell(lngRows, 3).Text) & ", " _
                            & NumericVal(Grid.Cell(lngRows, 4).Text) & ", " _
                            & NumericVal(Grid.Cell(lngRows, 5).Text) & ", " _
                            & NumericVal(Grid.Cell(lngRows, 6).Text) & ", " _
                            & NumericVal(Grid.Cell(lngRows, 7).Text) & ", " _
                            & NumericVal(Grid.Cell(lngRows, 8).Text) & ", " _
                            & NumericVal(Grid.Cell(lngRows, 9).Text) & ", " _
                            & NumericVal(Grid.Cell(lngRows, 10).Text) & ", " _
                            & NumericVal(Grid.Cell(lngRows, 11).Text) & ") "
                    gconDMIS.Execute TEMPSQL

                End With
        End Select
        lvwModules.ListItems.Remove (lvwModules.SelectedItem.Index)
    End If
End Sub
Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:
    lvwModules.Enabled = True
    Exit Sub
ErrorCode:
    ShowVBError
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdPrint_Click()
    Dim FILTER                          As String
    On Error GoTo ErrorCode:
    FILTER = " {U.USERID}=" & mUserID
    FILTER = FILTER & " AND {A.MAINMODULENAME}='" & cboMainModule & "'"
    rtpPrint.Reset
    rtpPrint.Formulas(0) = "CompanyName = '" & Company_name & "'"
    rtpPrint.Formulas(1) = "CompanyAddress = '" & Company_Address & "'"
    If UCase(cboModuleType) = "DATA ENTRY" Then
        PrintSQLReport rtpPrint, CRIS_REPORT_PATH & "\AccessReportDataEntry.rpt", FILTER, DMIS_REPORT_Connection, 1
    ElseIf UCase(cboModuleType) = "TRANSACTION" Then
        PrintSQLReport rtpPrint, CRIS_REPORT_PATH & "\AccessReportTransaction.rpt", FILTER, DMIS_REPORT_Connection, 1
    ElseIf UCase(cboModuleType) = "REPORTS" Then
        PrintSQLReport rtpPrint, CRIS_REPORT_PATH & "\AccessReportReport.rpt", FILTER, DMIS_REPORT_Connection, 1
    ElseIf UCase(cboModuleType) = "PROCESSING" Then
        PrintSQLReport rtpPrint, CRIS_REPORT_PATH & "\AccessReportProcessing.rpt", FILTER, DMIS_REPORT_Connection, 1
    ElseIf UCase(cboModuleType) = "INQUIRY" Then
        PrintSQLReport rtpPrint, CRIS_REPORT_PATH & "\AccessReportInquiry.rpt", FILTER, DMIS_REPORT_Connection, 1
    ElseIf UCase(cboModuleType) = "SYSTEM" Then
        PrintSQLReport rtpPrint, CRIS_REPORT_PATH & "\AccessReportSYSTEM.rpt", FILTER, DMIS_REPORT_Connection, 1
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub
Private Sub cmdSave_Click()
    On Error GoTo ErrorCode:
    Screen.MousePointer = 11
    DoEvents
    Call Save_Modules
    Screen.MousePointer = 0
    On Error GoTo ErrorCode
    gconDMIS.Execute ("DELETE FROM ALL_RAMS_USER_MODULES WHERE USERID NOT IN (SELECT USERID FROM ALL_RAMS_USERS) ")
    gconDMIS.Execute ("DELETE  FROM ALL_Rams_UsersAcess WHERE USERID NOT IN (SELECT USERID FROM ALL_RAMS_USERS)")
    gconDMIS.Execute ("DELETE  FROM ALL_Rams_UsersAcess WHERE MODULEID NOT IN (SELECT MODULEID  FROM ALL_RAMS_MODULES)")
    cmdSave.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub
Private Sub cboUsers_Click()
    If cboUsers.ListIndex = -1 Then
        Exit Sub
    End If
    mUserID = cboUsers.ItemData(cboUsers.ListIndex)
    NameStr = cboUsers.Text
    InitData
End Sub
Private Sub Command1_Click()
    On Error GoTo ErrorCode
    FillData
    Command1.Enabled = False
    cmdPrint.Enabled = True
    Check1.Value = 0
    Exit Sub
ErrorCode:
    ShowVBError
End Sub
Private Sub Form_Load()
    If CHANGE_USER = True Then
        Call FillCombo("SELECT  USER_NAME,USERID FROM ALL_RAMS_USERS WHERE LOCK=0 and USERGROUP<>'SDM' ORDER BY USER_NAME", 1, 0, cboUsers)
    Else
        Call FillCombo("SELECT  USERNAME,USERID FROM ALL_RAMS_USERS WHERE LOCK=0 and USERGROUP<>'SDM' ORDER BY USERNAME", 1, 0, cboUsers)
    End If
    InitGrid
    InitData
    If mUserID = 0 And cboUsers.ListCount > 0 Then
        cboUsers.ListIndex = 0
    ElseIf mUserID > 0 And cboUsers.ListCount > 0 Then
        cboUsers.ListIndex = SelectCombo(cboUsers, CStr(mUserID), True)
    Else
        MsgBox "There are No Users To Configure. Please Add Users First. Form Will Now Unload", vbInformation
        Unload Me
        Exit Sub
    End If
End Sub
Sub InitGrid()
    With Grid
        .AllowUserResizing = False
        .DisplayFocusRect = False
        .Appearance = Flat
        .ScrollBarStyle = Flat
        .FixedRowColStyle = Flat
        .BackColorFixed = RGB(90, 158, 214)
        .BackColorFixedSel = RGB(110, 180, 230)
        .BackColorBkg = RGB(90, 158, 214)
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)
    End With
End Sub
Sub InitData()
    Call FillCombo("SELECT  MAINMODULENAME  FROM ALL_RAMS_USER_MODULES where USERID=" & mUserID, -1, 0, cboMainModule)
    Call Combo_Loadval(cboModuleType, gconDMIS.Execute("SELECT DISTINCT UPPER(MODULE_TYPE) from ALL_RAMS_MODULES WHERE MODULE_TYPE <>'SEARCH' OR MODULE_TYPE IS NOT NULL"))

    If cboMainModule.ListCount > 0 Then
        cboMainModule.ListIndex = 0
        cboModuleType.ListIndex = 0
    End If
    cmdPrint.Enabled = False
    cmdSave.Enabled = False
    picProgress.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If cmdSave.Enabled = False Then: Exit Sub
    If MsgBox("Record Has Been Changed. Do You Want to Exit without Saving?", vbInformation + vbYesNo) = vbNo Then
    Cancel = 1
    End If
End Sub

Private Sub Grid_Click()
If Grid.ActiveCell.Row = 0 Then Exit Sub
    On Error GoTo adder
    Dim gcol                            As Long
    gcol = Grid.Selection.FirstCol
    If gcol <> Grid.Cols - 1 And gcol <> 0 And gcol <> 1 Then
        If Grid.Cell(Grid.Selection.FirstRow, gcol).Text = "" Then
            Grid.Cell(Grid.Selection.FirstRow, gcol).Text = 1
        ElseIf Grid.Cell(Grid.Selection.FirstRow, gcol).Text = 0 Then
            Grid.Cell(Grid.Selection.FirstRow, gcol).Text = 1
        Else
            Grid.Cell(Grid.Selection.FirstRow, gcol).Text = 0
        End If
        cmdSave.Enabled = True
    End If
    Grid.Refresh
    Exit Sub
adder:
    Err.Clear
End Sub
Private Sub Grid_DblClick()
If Grid.ActiveCell.Row = 0 Then Exit Sub
    If Grid.Selection.FirstCol = Grid.Cols - 1 Then
        Dim LST                         As ListItem
        gconDMIS.Execute "Delete From ALL_RAMS_USERSACESS WHERE MODULEID =" & Grid.Cell(Grid.Selection.FirstRow, 0).Text & " AND USERID=" & mUserID
        Set LST = lvwModules.ListItems.Add(, , Grid.Cell(Grid.Selection.FirstRow, 1).Text)
        LST.ListSubItems.Add , , Grid.Cell(Grid.Selection.FirstRow, 0).Text
        Grid.RemoveItem (Grid.Selection.FirstRow)
        Set LST = Nothing
    End If
End Sub

Private Sub Grid_EditRow(ByVal Row As Long)
    cmdSave.Enabled = True
End Sub
 

Private Sub lvwModules_DblClick()
    If lvwModules.SelectedItem Is Nothing Then Exit Sub
    cmdAdd_Click
End Sub
Private Sub mnuExportSettings_Click()
    If Grid.ExportToXML("") Then
        MsgBox "OK", vbExclamation
    End If
End Sub
Private Sub mnuImportSettings_Click()
    If Grid.LoadFromXML("") Then
        MsgBox "OK"
    End If
End Sub

Private Sub Text1_Change()
    Dim NotInSQL                        As String
    Dim RS                              As ADODB.Recordset
    If LTrim(RTrim(Text1.Text)) = "" Then
        NotInSQL = "SELECT   DESCRIPTIONS,  MODULEID  FROM ALL_RAMS_MODULES " & _
                 " WHERE    MODULEID NOT IN(SELECT MODULEID FROM ALL_vW_RAMS_USERACESS  WHERE userid = " & mUserID & ") " & _
                 "  and  MAINMODULENAME = '" & cboMainModule.Text & "' and MODULE_TYPE='" & UCase(cboModuleType.Text) & "' Order by DESCRIPTIONS"

    Else
        NotInSQL = "SELECT   DESCRIPTIONS,  MODULEID  FROM ALL_RAMS_MODULES " & _
                 " WHERE    MODULEID NOT IN(SELECT MODULEID FROM ALL_vW_RAMS_USERACESS  WHERE userid = " & mUserID & ") " & _
                 "  and  MAINMODULENAME = '" & cboMainModule.Text & "' and MODULE_TYPE='" & UCase(cboModuleType.Text) & "' AND DESCRIPTIONS LIKE '%" & Repleys(Text1) & "%'  Order by DESCRIPTIONS"
    End If
    Set RS = gconDMIS.Execute(NotInSQL)
    If Not (RS.BOF And RS.EOF) Then
        lvwModules.Enabled = True
        Listview_Loadval Me.lvwModules.ListItems, RS
    Else
        lvwModules.Enabled = False
        Me.lvwModules.ListItems.Clear
    End If

    Set RS = Nothing
End Sub
    '    SQL = " Alter table ALL_Rams_UsersAcess " & vbCrLf
    '    SQL = SQL & "  DROP column [ID] " & vbCrLf
    '    SQL = SQL & " Alter table ALL_Rams_UsersAcess " & vbCrLf
    '    SQL = SQL & " ADD  [ID] [int] IDENTITY (1, 1) NOT NULL "


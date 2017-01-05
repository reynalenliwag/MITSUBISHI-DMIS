VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "WIZFLEX.OCX"
Begin VB.Form frmUserModules 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " User Modules"
   ClientHeight    =   8985
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   14460
   ClipControls    =   0   'False
   ForeColor       =   &H00F5F5F5&
   Icon            =   "UserModules.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   14460
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   13680
      MouseIcon       =   "UserModules.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "UserModules.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8220
      Width           =   705
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   12930
      MouseIcon       =   "UserModules.frx":12FA
      MousePointer    =   99  'Custom
      Picture         =   "UserModules.frx":144C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8220
      Width           =   705
   End
   Begin VB.ComboBox cboModuleType 
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
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   390
      Width           =   3525
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
      Left            =   12210
      TabIndex        =   9
      Top             =   1440
      Width           =   1785
   End
   Begin VB.ComboBox cboMainModule 
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
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   390
      Width           =   3525
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   12180
      MouseIcon       =   "UserModules.frx":178A
      MousePointer    =   99  'Custom
      Picture         =   "UserModules.frx":18DC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8220
      Width           =   705
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
      Height          =   435
      Left            =   2850
      MouseIcon       =   "UserModules.frx":1C2C
      MousePointer    =   99  'Custom
      Picture         =   "UserModules.frx":1D7E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1290
      Width           =   945
   End
   Begin MSComctlLib.ListView lvwModules 
      Height          =   6960
      Left            =   60
      TabIndex        =   5
      Top             =   1710
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   12277
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin FlexCell.Grid Grid 
      Height          =   6435
      Left            =   3480
      TabIndex        =   10
      Top             =   1740
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   11351
      Appearance      =   0
      BackColorBkg    =   -2147483645
      BackColorSel    =   16777215
      Cols            =   6
      DefaultFontName =   "Arial"
      DefaultFontSize =   8.25
      DisplayRowIndex =   -1  'True
      Rows            =   1
      ScrollBarStyle  =   0
      SelectionMode   =   1
      EnterKeyMoveTo  =   1
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Module Name"
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
      Left            =   3750
      TabIndex        =   16
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System"
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
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE PRESS UPDATE AFTER EDITING CELL"
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
      Left            =   10200
      TabIndex        =   12
      Top             =   30
      Width           =   3780
   End
   Begin VB.Label lblDescription 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7320
      TabIndex        =   8
      Top             =   360
      Width           =   6675
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
      Left            =   4890
      TabIndex        =   3
      Top             =   1440
      Width           =   2670
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username :"
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
      TabIndex        =   0
      Top             =   945
      Width           =   975
   End
   Begin VB.Label lblUsername 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   1335
      TabIndex        =   1
      Top             =   945
      Width           =   5475
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available Modules To be Added:"
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
      Left            =   180
      TabIndex        =   2
      Top             =   1380
      Width           =   2625
   End
End
Attribute VB_Name = "frmUserModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NameStr                                  As String
Dim itemInt                                  As Integer
Dim mUserID                                  As Long
Dim SQL                                      As String
Dim I                                        As Integer
Dim TempRs                                   As ADODB.Recordset
Public Property Let UserID(sUserID As Long)
    mUserID = sUserID
End Property

Public Property Let Username(sNamestr As String)
    NameStr = sNamestr
End Property

Private Sub cboMainModule_CLICK()
    cboModuleType_Click
End Sub

Private Sub cboModuleType_Click()
    If cboModuleType.ListIndex = -1 Then: Exit Sub
    Dim rs                                   As ADODB.Recordset
    Dim NotInSQL                             As String
    NotInSQL = "SELECT   DESCRIPTIONS,  ID  FROM ALL_RAMS_MODULES " & _
             " WHERE    ID NOT IN(SELECT MODULEID FROM ALL_vW_RAMS_USERACESS  WHERE userid = " & mUserID & ") " & _
             "  and  MAINMODULEID = " & cboMainModule.ItemData(cboMainModule.ListIndex) & " and MODULE_TYPE='" & UCase(cboModuleType.Text) & "'"
    Set rs = gconDMIS.Execute(NotInSQL)

    If Not (rs.BOF And rs.EOF) Then
        lvwModules.Enabled = True
        Listview_Loadval Me.lvwModules.ListItems, rs
    Else
        lvwModules.Enabled = False
        Me.lvwModules.ListItems.Clear
    End If

    Set rs = Nothing

    Screen.MousePointer = 11
    I = 0
    Grid.Visible = False
    Select Case UCase(cboModuleType.Text)
        Case "SYSTEM"
            ShowACCESS_System cboMainModule.ItemData(cboMainModule.ListIndex), mUserID, cboModuleType.Text
        Case "DATA ENTRY"
            ShowACCESS_DataEntry cboMainModule.ItemData(cboMainModule.ListIndex), mUserID, cboModuleType.Text
        Case "SEARCH"
            ShowACCESS_SEARCH cboMainModule.ItemData(cboMainModule.ListIndex), mUserID, cboModuleType.Text
        Case "INQUIRY"
            ShowACCESS_INQUIRY cboMainModule.ItemData(cboMainModule.ListIndex), mUserID, cboModuleType.Text
        Case "REPORTS"
            ShowACCESS_Reports cboMainModule.ItemData(cboMainModule.ListIndex), mUserID, cboModuleType.Text
        Case "PROCESSING"
            ShowACCESS_Processing cboMainModule.ItemData(cboMainModule.ListIndex), mUserID, cboModuleType.Text
        Case "TRANSACTION"
            ShowACCESS_TRANSACTION cboMainModule.ItemData(cboMainModule.ListIndex), mUserID, cboModuleType.Text
    End Select
    Grid.Visible = True
    Set TempRs = Nothing
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
        .Column(colINDEX).Locked = True
        .Column(colINDEX).CellType = cellHyperLink
        .Column(colINDEX).Locked = True
        .Column(colINDEX).Width = 50
        .Cell(0, colINDEX).Text = "OPTIONS"
    End With
End Sub

Sub ShowACCESS_System(XMainModuleID, XUserID, XmoduleType)
    SQL = "SELECT ARU.MODULEID, ARM.DESCRIPTIONS,  ARU.Acess_System  " & vbCrLf & _
        " FROM ALL_RAMS_USERSACESS ARU INNER JOIN  ALL_RAMS_MODULES  ARM ON ARU.MODULEID = ARM.ID " & vbCrLf & _
        " WHERE (ARM.MAINMODULEID = " & XMainModuleID & " AND ARU.USERID = " & XUserID & " AND ARM.MODULE_TYPE = '" & XmoduleType & "') "
    With Grid
        .Cols = 4
        .FixedCols = 1
        .Column(1).Width = 281
        .Column(2).Width = 75
        .Cell(0, 1).Text = "MODULE NAME"
        .Cell(0, 2).Text = "ACCESS"

        .Column(2).CellType = cellCheckBox
        Call AddHyperLink(3)
    End With

    Set TempRs = gconDMIS.Execute(SQL)
    Grid.Rows = 1
    While Not TempRs.EOF
        With Grid
            I = I + 1
            .AddItem TempRs!Descriptions, False
            .Cell(I, 0).Text = Null2String(TempRs!ModuleID)
            .Cell(I, 2).Text = Null2String(TempRs!Acess_System)
            .Cell(I, 3).Font.Bold = True
            .Cell(I, 3).ForeColor = vbBlue
            .Cell(I, 3).Text = "DELETE"
        End With
        TempRs.MoveNext
    Wend
    Grid.Refresh

End Sub


Sub ShowACCESS_DataEntry(XMainModuleID, XUserID, XmoduleType)

    SQL = "SELECT ARU.MODULEID, ARM.DESCRIPTIONS, ARU.Acess_Add, ARU.Acess_Edit, ARU.Acess_Delete, ARU.Acess_View, ARU.Acess_Print, ARU.Acess_Process  " & vbCrLf & _
        " FROM ALL_RAMS_USERSACESS ARU INNER JOIN  ALL_RAMS_MODULES  ARM ON ARU.MODULEID = ARM.ID " & vbCrLf & _
        " WHERE (ARM.MAINMODULEID = " & XMainModuleID & " AND ARU.USERID = " & XUserID & " AND ARM.MODULE_TYPE='" & XmoduleType & "' ) "
    With Grid
        .Cols = 8
        .Column(1).Width = 190
        .Column(2).Width = 75
        .Column(3).Width = 75
        .Column(4).Width = 75
        .Column(5).Width = 75
        .Column(6).Width = 75
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
    Set TempRs = gconDMIS.Execute(SQL)

    Grid.Rows = 1
    While Not TempRs.EOF
        With Grid
            I = I + 1
            .AddItem TempRs!Descriptions, False
            .Cell(I, 7).ForeColor = vbBlue
            .Cell(I, 7).Text = "DELETE"
            .Cell(I, 7).Font.Bold = True
            .Cell(I, 0).Text = (TempRs!ModuleID)
            .Cell(I, 2).Text = (TempRs!Acess_Add)
            .Cell(I, 3).Text = Null2String(TempRs!Acess_Edit)
            .Cell(I, 4).Text = Null2String(TempRs!Acess_View)
            .Cell(I, 5).Text = Null2String(TempRs!Acess_Delete)
            .Cell(I, 6).Text = Null2String(TempRs!Acess_Print)
        End With
        TempRs.MoveNext
    Wend

    Grid.Refresh

End Sub
Sub ShowACCESS_SEARCH(XMainModuleID, XUserID, XmoduleType)
    SQL = "SELECT ARU.MODULEID, ARM.DESCRIPTIONS, ARU.Acess_System " & vbCrLf & _
        " FROM ALL_RAMS_USERSACESS ARU INNER JOIN  ALL_RAMS_MODULES  ARM ON ARU.MODULEID = ARM.ID " & vbCrLf & _
        " WHERE (ARM.MAINMODULEID = " & XMainModuleID & " AND ARU.USERID = " & XUserID & " AND ARM.MODULE_TYPE='" & XmoduleType & "' ) "
    With Grid
        .Cols = 4
        .Column(1).Width = 281
        .Column(2).Width = 75
        .Column(2).CellType = cellCheckBox
        AddHyperLink (3)
        .Cell(0, 1).Text = "MODULE NAME"
        .Cell(0, 2).Text = "ACCESS"

    End With
    Set TempRs = gconDMIS.Execute(SQL)

    Grid.Rows = 1
    While Not TempRs.EOF
        With Grid
            I = I + 1
            .AddItem TempRs!Descriptions, False
            .Cell(I, 0).Text = Null2String(TempRs!ModuleID)
            .Cell(I, 2).Text = Null2String(TempRs!Acess_System)
            .Cell(I, 3).ForeColor = vbBlue
            .Cell(I, 3).Text = "DELETE"
            .Cell(I, 3).Font.Bold = True
        End With

        TempRs.MoveNext
    Wend

    Grid.Refresh
End Sub


Sub ShowACCESS_INQUIRY(XMainModuleID, XUserID, XmoduleType)
    SQL = "SELECT ARU.MODULEID, ARM.DESCRIPTIONS, ARU.Acess_System " & vbCrLf & _
        " FROM ALL_RAMS_USERSACESS ARU INNER JOIN  ALL_RAMS_MODULES  ARM ON ARU.MODULEID = ARM.ID " & vbCrLf & _
        " WHERE (ARM.MAINMODULEID = " & XMainModuleID & " AND ARU.USERID = " & XUserID & " AND ARM.MODULE_TYPE='" & XmoduleType & "' ) "
    With Grid
        .Cols = 3
        .Column(1).Width = 281
        .Column(2).Width = 75
        .Column(2).CellType = cellCheckBox
        .Cell(0, 1).Text = "MODULE NAME"
        .Cell(0, 2).Text = "ACCESS"
        AddHyperLink (2)
    End With
    Set TempRs = gconDMIS.Execute(SQL)

    Grid.Rows = 1
    While Not TempRs.EOF
        With Grid
            I = I + 1
            .AddItem TempRs!Descriptions, False
            .Cell(I, 0).Text = Null2String(TempRs!ModuleID)
            .Cell(I, 2).Text = Null2String(TempRs!Acess_System)
            .Cell(I, 3).ForeColor = vbBlue
            .Cell(I, 3).Text = "DELETE"
            .Cell(I, 3).Font.Bold = True
        End With

        TempRs.MoveNext
    Wend

    Grid.Refresh
End Sub

Sub ShowACCESS_Reports(XMainModuleID, XUserID, XmoduleType)
    SQL = "SELECT ARU.MODULEID, ARM.DESCRIPTIONS, ARU.Acess_View, ARU.Acess_Print  " & vbCrLf & _
        " FROM ALL_RAMS_USERSACESS ARU INNER JOIN  ALL_RAMS_MODULES  ARM ON ARU.MODULEID = ARM.ID " & vbCrLf & _
        " WHERE (ARM.MAINMODULEID = " & XMainModuleID & " AND ARU.USERID = " & XUserID & " AND ARM.MODULE_TYPE = '" & XmoduleType & "') "
    With Grid
        .Cols = 5
        .Column(1).Width = 281
        .Column(2).Width = 75
        .Column(3).Width = 75
        .Column(2).CellType = cellCheckBox
        .Column(3).CellType = cellCheckBox
        .Cell(0, 1).Text = "MODULE NAME"
        .Cell(0, 2).Text = "VIEW"
        .Cell(0, 3).Text = "PRINT"
        AddHyperLink (4)
    End With
    Set TempRs = gconDMIS.Execute(SQL)

    Grid.Rows = 1
    While Not TempRs.EOF
        With Grid
            I = I + 1
            .AddItem TempRs!Descriptions, False
            .Cell(I, 0).Text = Null2String(TempRs!ModuleID)
            .Cell(I, 2).Text = Null2String(TempRs!Acess_View)
            .Cell(I, 3).Text = Null2String(TempRs!Acess_Print)
            .Cell(I, 4).ForeColor = vbBlue
            .Cell(I, 4).Text = "DELETE"
            .Cell(I, 4).Font.Bold = True
        End With
        TempRs.MoveNext
    Wend

    Grid.Refresh
End Sub

Sub ShowACCESS_Processing(XMainModuleID, XUserID, XmoduleType)
    SQL = "SELECT ARU.MODULEID, ARM.DESCRIPTIONS, ARU.Acess_Process" & vbCrLf & _
        " FROM ALL_RAMS_USERSACESS ARU INNER JOIN  ALL_RAMS_MODULES  ARM ON ARU.MODULEID = ARM.ID " & vbCrLf & _
        " WHERE (ARM.MAINMODULEID = " & XMainModuleID & " AND ARU.USERID = " & XUserID & " AND ARM.MODULE_TYPE = '" & XmoduleType & "') "

    With Grid
        .Cols = 4
        .Column(1).Width = 281
        .Column(2).Width = 150
        .Column(2).CellType = cellCheckBox
        AddHyperLink (3)
        .Cell(0, 1).Text = "MODULE NAME"
        .Cell(0, 2).Text = "IMPORT/PROCESS"


    End With
    Set TempRs = gconDMIS.Execute(SQL)

    Grid.Rows = 1
    While Not TempRs.EOF
        With Grid
            I = I + 1
            .AddItem TempRs!Descriptions, False
            .Cell(I, 0).Text = Null2String(TempRs!ModuleID)
            .Cell(I, 2).Text = Null2String(TempRs!Acess_PROCESS)
            .Cell(I, 3).Text = "DELETE"
            .Cell(I, 3).ForeColor = vbBlue
            .Cell(I, 3).Font.Bold = True
        End With
        TempRs.MoveNext
    Wend

    Grid.Refresh
End Sub
Sub ShowACCESS_TRANSACTION(XMainModuleID, XUserID, XmoduleType)
    SQL = "SELECT ARU.MODULEID, ARM.DESCRIPTIONS, ARU.Acess_Add, ARU.Acess_Edit, ARU.Acess_Delete, ARU.Acess_CancelEntry, ARU.Acess_Print, ARU.Acess_Post, ARU.Acess_UnPost " & vbCrLf & _
        " FROM ALL_RAMS_USERSACESS ARU INNER JOIN  ALL_RAMS_MODULES  ARM ON ARU.MODULEID = ARM.ID " & vbCrLf & _
        " WHERE (ARM.MAINMODULEID = " & XMainModuleID & " AND ARU.USERID = " & XUserID & " AND ARM.MODULE_TYPE = '" & XmoduleType & "') "

    With Grid
        .Cols = 10
        .Column(1).Width = 281
        .Column(2).Width = 75
        .Column(3).Width = 75
        .Column(4).Width = 75
        .Column(5).Width = 75
        .Column(6).Width = 75
        .Column(7).Width = 75
        .Column(8).Width = 75

        .Column(2).CellType = cellCheckBox
        .Column(3).CellType = cellCheckBox
        .Column(4).CellType = cellCheckBox
        .Column(5).CellType = cellCheckBox
        .Column(6).CellType = cellCheckBox
        .Column(7).CellType = cellCheckBox
        .Column(8).CellType = cellCheckBox

        AddHyperLink (9)
        .Cell(0, 1).Text = "MODULE NAME"
        .Cell(0, 2).Text = "ADD"
        .Cell(0, 3).Text = "EDIT"
        .Cell(0, 5).Text = "DELETE"
        .Cell(0, 4).Text = "PRINT"
        .Cell(0, 6).Text = "POST"
        .Cell(0, 7).Text = "UN-POST"
        .Cell(0, 8).Text = "CANCEL ENTRY"


    End With
    Set TempRs = gconDMIS.Execute(SQL)

    Grid.Rows = 1
    While Not TempRs.EOF
        With Grid
            I = I + 1
            .AddItem TempRs!Descriptions, False
            .Cell(I, 0).Text = Null2String(TempRs!ModuleID)
            .Cell(I, 2).Text = Null2String(TempRs!Acess_Add)
            .Cell(I, 3).Text = Null2String(TempRs!Acess_Edit)
            .Cell(I, 4).Text = Null2String(TempRs!Acess_Delete)
            .Cell(I, 5).Text = Null2String(TempRs!Acess_Print)
            .Cell(I, 6).Text = Null2String(TempRs!Acess_POST)
            .Cell(I, 7).Text = Null2String(TempRs!Acess_UnPost)
            .Cell(I, 8).Text = Null2String(TempRs!Acess_CancelEntry)
            .Cell(I, 9).ForeColor = vbBlue
            .Cell(I, 9).Text = "DELETE"
            .Cell(I, 9).Font.Bold = True
        End With
        TempRs.MoveNext
    Wend

    Grid.Refresh

End Sub



Private Sub Save_Modules()
    On Error GoTo CNERRORS
    Dim TEMPSQL                              As String
    Dim lxJ                                  As Long

    gconDMIS.BeginTrans
    gconDMIS.Execute ("Delete from ALL_RAMS_USERSACESS where MODULEID IN (Select MODULEID from ALL_vW_RAMS_USERACESS WHERE  MODULE_TYPE='" & UCase(cboModuleType.Text) & "' AND USERID=" & mUserID & ")")

    Select Case UCase(cboModuleType.Text)
        Case "DATA ENTRY"
            For lxJ = 1 To Grid.Rows - 1
                TEMPSQL = " INSERT INTO ALL_vW_RAMS_USERACESS " & _
                        " ( USERID, MODULEID , Acess_Add, Acess_Edit, Acess_View, Acess_Delete, Acess_Print) " & _
                        " VALUES(" & mUserID & ", " & Grid.Cell(lxJ, 0).Text & ", " & Grid.Cell(lxJ, 2).Text & ", 1 , 1, 1 ,1) "
                gconDMIS.Execute TEMPSQL

            Next
        Case "SEARCH", "SYSTEM", "INQUIRY"
            For lxJ = 1 To Grid.Rows - 1
                TEMPSQL = " INSERT INTO ALL_vW_RAMS_USERACESS " & _
                        " ( USERID, MODULEID , Acess_System) " & _
                        " VALUES(" & mUserID & ", " & Grid.Cell(lxJ, 0).Text & ", " & Grid.Cell(lxJ, 2).Text & ") "
                gconDMIS.Execute TEMPSQL
            Next
        Case "REPORTS"
            For lxJ = 1 To Grid.Rows - 1
                TEMPSQL = " INSERT INTO ALL_vW_RAMS_USERACESS " & _
                        " ( USERID, MODULEID , Acess_View, Acess_Print) " & _
                        " VALUES(" & mUserID & ", " & Grid.Cell(lxJ, 0).Text & ", " & Grid.Cell(lxJ, 2).Text & ", " & Grid.Cell(lxJ, 3).Text & ") "
                gconDMIS.Execute TEMPSQL

            Next
        Case "PROCESSING"
            For lxJ = 1 To Grid.Rows - 1
                TEMPSQL = " INSERT INTO ALL_vW_RAMS_USERACESS " & _
                        " ( USERID, MODULEID , Acess_Process) " & _
                        " VALUES(" & mUserID & ", " & Grid.Cell(lxJ, 0).Text & ", " & Grid.Cell(lxJ, 2).Text & " ) "
                gconDMIS.Execute TEMPSQL

            Next
        Case "TRANSACTION"
            For lxJ = 1 To Grid.Rows - 1
                TEMPSQL = " INSERT INTO ALL_vW_RAMS_USERACESS " & _
                        " ( USERID, MODULEID , Acess_Add, Acess_Edit, Acess_Print ,  Acess_Delete, Acess_POST, Acess_UnPost, Acess_CancelEntry ) " & _
                        " VALUES(" & mUserID & ", " _
                        & Grid.Cell(lxJ, 0).Text & ", " _
                        & Grid.Cell(lxJ, 2).Text & ", " _
                        & Grid.Cell(lxJ, 3).Text & ", " _
                        & Grid.Cell(lxJ, 4).Text & ", " _
                        & Grid.Cell(lxJ, 5).Text & ", " _
                        & Grid.Cell(lxJ, 6).Text & ", " _
                        & Grid.Cell(lxJ, 7).Text & ", " _
                        & Grid.Cell(lxJ, 8).Text & ") "
                gconDMIS.Execute TEMPSQL

            Next
    End Select

    gconDMIS.CommitTrans

    Exit Sub
CNERRORS:
    gconDMIS.RollbackTrans
    MsgBox Err.Description
    Err.Clear

    MsgBox "Cannot Process your Request.."
End Sub



Private Sub Check1_Click()
    Dim j                                    As Long

    For I = 2 To Grid.Cols - 2
        For j = 1 To Grid.Rows - 1
            Grid.Cell(j, I).Text = Check1.Value
        Next

    Next


    Grid.Refresh
End Sub

Private Sub cmdAdd_Click()
    If lvwModules.SelectedItem Is Nothing Then
        MessagePop InfoVoid, "Selection Required", "There is Nothing To Select"
        Exit Sub

    End If

    If Me.lvwModules.ListItems.Count > 0 Then
        Dim lngRows                          As Long
        
        gconDMIS.Execute ("INSERT INTO ALL_RAMS_USERSACESS ( MODULEID, USERID) values(" & lvwModules.SelectedItem.ListSubItems(1).Text & "," & mUserID & ")")
    
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
            Case "PROCESSING"
                With Grid
                    .AddItem lvwModules.SelectedItem.Text, True
                    .Cell(I, 0).Text = lvwModules.SelectedItem.ListSubItems(1).Text
                    .Cell(I, 2).Text = 1
                    .Cell(I, 3).ForeColor = vbBlue
                    .Cell(I, 3).Text = "DELETE"
                    .Cell(I, 3).Font.Bold = True
                End With
            Case "TRANSACTION"
                With Grid
                    .AddItem lvwModules.SelectedItem.Text, True
                    .Cell(I, 0).Text = lvwModules.SelectedItem.ListSubItems(1).Text
                    .Cell(I, 2).Text = 1
                    .Cell(I, 3).Text = 1
                    .Cell(I, 4).Text = 1
                    .Cell(I, 5).Text = 1
                    .Cell(I, 6).Text = 1
                    .Cell(I, 7).Text = 1
                    .Cell(I, 8).Text = 1
                    .Cell(I, 9).ForeColor = vbBlue
                    .Cell(I, 9).Text = "DELETE"
                    .Cell(I, 9).Font.Bold = True
                End With
        End Select
        lvwModules.ListItems.Remove (lvwModules.SelectedItem.Index)
    End If
End Sub


Private Sub cmdCancel_Click()
    cboMainModule.Enabled = True
    cboModuleType.Enabled = True

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()


    Call Save_Modules
    cboModuleType.Enabled = True
    cboMainModule.Enabled = True
End Sub


Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Call FillCombo("SELECT  AP.ModuleName, AP.ID FROM ALL_RAMS_USER_MODULES ARUM INNER JOIN ALL_Profile AP ON ARUM.MAINMODULEID = AP.ID WHERE (ARUM.USERID = " & mUserID & ")", 1, 0, cboMainModule)
    Call Combo_Loadval(cboModuleType, gconDMIS.Execute("SELECT DISTINCT UPPER(MODULE_TYPE) from ALL_RAMS_MODULES WHERE MODULE_TYPE IS NOT NULL"))
    cboMainModule.ListIndex = 0
    cboModuleType.ListIndex = 0

    Me.lblUsername.Caption = NameStr
    Grid.Column(1).Locked = True

End Sub


Private Sub Grid_DblClick()
    If Grid.Selection.FirstCol = Grid.Cols - 1 Then

        Dim LST                              As ListItem
        gconDMIS.Execute "Delete From ALL_RAMS_USERSACESS WHERE MODULEID =" & Grid.Cell(Grid.Selection.FirstRow, 0).Text & " AND USERID=" & mUserID
        Set LST = lvwModules.ListItems.Add(, , Grid.Cell(Grid.Selection.FirstRow, 1).Text)
        LST.ListSubItems.Add , , Grid.Cell(Grid.Selection.FirstRow, 0).Text
        Grid.RemoveItem (Grid.Selection.FirstRow)
        Set LST = Nothing

    End If
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmRAMS_AEModule 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modules"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9045
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AEModule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDetails 
      Height          =   5205
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   2745
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   90
         MaxLength       =   35
         TabIndex        =   1
         Top             =   240
         Width           =   2565
      End
      Begin MSComctlLib.ListView lstGrid 
         Height          =   4455
         Left            =   60
         TabIndex        =   2
         Top             =   660
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   7858
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
         MouseIcon       =   "AEModule.frx":000C
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Module Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   2760
      TabIndex        =   4
      Top             =   -30
      Width           =   6255
      Begin VB.ComboBox cboMainModuleName 
         BackColor       =   &H00E0E0E0&
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
         Height          =   345
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "cboMainModuleName"
         Top             =   720
         Width           =   3345
      End
      Begin VB.TextBox txtModuleName 
         BackColor       =   &H00E0E0E0&
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
         Height          =   1275
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   185
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "AEModule.frx":016E
         Top             =   1350
         Width           =   3435
      End
      Begin VB.ComboBox cboModuleType 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3300
         Width           =   2655
      End
      Begin MSComctlLib.ImageList ImageList 
         Left            =   5190
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   160
         ImageHeight     =   94
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AEModule.frx":0174
               Key             =   "AMIS"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AEModule.frx":2694
               Key             =   "CMIS"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AEModule.frx":46F1
               Key             =   "CRIS"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AEModule.frx":6E48
               Key             =   "CSMS"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AEModule.frx":8D4C
               Key             =   "HRMS"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AEModule.frx":B0B8
               Key             =   "PMIS"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AEModule.frx":D329
               Key             =   "SMIS"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "AEModule.frx":F5D1
               Key             =   "WMIS"
            EndProperty
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   1305
         Left            =   3720
         Top             =   780
         Width           =   2265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Main Module"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   180
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Module Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Left            =   180
         TabIndex        =   8
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Module Type (* Editable Fields)"
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
         Left            =   210
         TabIndex        =   11
         Top             =   2970
         Width           =   2565
      End
      Begin XtremeShortcutBar.ShortcutCaption CapInfo 
         Height          =   270
         Index           =   2
         Left            =   0
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         Width           =   6240
         _Version        =   655364
         _ExtentX        =   11007
         _ExtentY        =   476
         _StockProps     =   14
         Caption         =   "Edit Module"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         ForeColor       =   64
      End
      Begin VB.Label labid 
         Alignment       =   2  'Center
         BackColor       =   &H00404000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3870
         TabIndex        =   9
         Top             =   1200
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         BorderStyle     =   6  'Inside Solid
         Height          =   1425
         Left            =   3660
         Top             =   720
         Width           =   2145
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   2190
      ScaleHeight     =   900
      ScaleWidth      =   7155
      TabIndex        =   13
      Top             =   4320
      Width           =   7155
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   6120
         MouseIcon       =   "AEModule.frx":11A8C
         MousePointer    =   99  'Custom
         Picture         =   "AEModule.frx":11BDE
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   5430
         MouseIcon       =   "AEModule.frx":11F44
         MousePointer    =   99  'Custom
         Picture         =   "AEModule.frx":12096
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   4740
         MouseIcon       =   "AEModule.frx":123F2
         MousePointer    =   99  'Custom
         Picture         =   "AEModule.frx":12544
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   4050
         MouseIcon       =   "AEModule.frx":1283E
         MousePointer    =   99  'Custom
         Picture         =   "AEModule.frx":12990
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   3360
         MouseIcon       =   "AEModule.frx":12CE8
         MousePointer    =   99  'Custom
         Picture         =   "AEModule.frx":12E3A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   2670
         MouseIcon       =   "AEModule.frx":13199
         MousePointer    =   99  'Custom
         Picture         =   "AEModule.frx":132EB
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7530
      ScaleHeight     =   885
      ScaleWidth      =   2580
      TabIndex        =   20
      Top             =   4320
      Width           =   2580
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   780
         MouseIcon       =   "AEModule.frx":135FE
         MousePointer    =   99  'Custom
         Picture         =   "AEModule.frx":13750
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   90
         MouseIcon       =   "AEModule.frx":13A8E
         MousePointer    =   99  'Custom
         Picture         =   "AEModule.frx":13BE0
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   540
      TabIndex        =   3
      Top             =   285
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmRAMS_AEModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsModules                                          As ADODB.Recordset
Dim AddorEdit                                          As String
Dim ModuleID                                           As Long
Event ChangedRecord(o As Boolean)

Private Sub cboMainModuleName_Change()
    On Error GoTo adder:
    Image1.Picture = ImageList.ListImages(UCase(cboMainModuleName.Text)).Picture
    Shape1.Move Image1.Left - 60, Image1.Top - 60, Image1.Width + 120, Image1.Height + 120
    Exit Sub
adder:
    Image1.Picture = ImageList.ListImages(1).Picture
End Sub

Private Sub cboMainModuleName_Click()
    cboMainModuleName_Change
End Sub

Private Sub cmdAdd_Click()
    AddorEdit = "ADD"
    cboMainModuleName.Locked = False
    cboMainModuleName.BackColor = vbWhite
    txtModuleName.Locked = False
    txtModuleName.BackColor = vbWhite
    Picture1.Visible = False
    Picture2.Visible = True
    Frame1.Enabled = True
    InitMemVars
    cboMainModuleName.SetFocus

End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    cboMainModuleName.Locked = True
    cboMainModuleName.BackColor = vbButtonFace
    txtModuleName.Locked = True
    txtModuleName.BackColor = vbButtonFace
    Picture1.Visible = True
    Picture2.Visible = False
    StoreMemvars
End Sub



'Upating Code       : AXP-0713200715:23
Private Sub cmdEdit_Click()
On Error GoTo Errorcode:

    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
    cboModuleType.SetFocus
Exit Sub
Errorcode:
ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    txtSearch.SetFocus
End Sub



'Upating Code       : AXP-0713200715:23
Private Sub cmdNext_Click()
On Error GoTo Errorcode:

    rsModules.MoveNext
    If rsModules.EOF Then
        rsModules.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
Exit Sub
Errorcode:
ShowVBError
End Sub

'Upating Code       : AXP-0713200715:23
Private Sub cmdPrevious_Click()
On Error GoTo Errorcode:

    rsModules.MovePrevious
    If rsModules.BOF Then
        rsModules.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
Exit Sub
Errorcode:
ShowVBError
End Sub

'Upating Code       : AXP-0713200715:23
Private Sub cmdSave_Click()

    Dim vtxtModeType                                   As String
    Dim vtxtMainModuleID                               As Long
    Dim vtxtModuleName                                 As String

On Error GoTo Errorcode:

    vtxtModeType = N2Str2Null(UCase(cboModuleType.Text))
    vtxtModuleName = N2Str2Null(UCase(txtModuleName.Text))


    If AddorEdit = "ADD" Then
        'LogAudit "A", "USER MODULES"
        cboMainModuleName.ListIndex = SelectCombo(cboMainModuleName, cboMainModuleName.Text)

        If cboMainModuleName.ListIndex = -1 Then: cboMainModuleName.SetFocus: Exit Sub
        vtxtMainModuleID = cboMainModuleName.ItemData(cboMainModuleName.ListIndex)
        If vtxtModuleName = "" Then: cboMainModuleName.SetFocus: Exit Sub
        gconDMIS.Execute "INSERT INTO  ALL_RAMS_MODULES (MainModuleID,Descriptions,Module_Type)" & _
                       " values(" & vtxtMainModuleID & " ," & vtxtModuleName & " ," & vtxtModeType & " )"


        MessagePop RecSaveOk, "New Module Added", " New Module Information Sucessfully Added For " & cboMainModuleName.Text & " System "
        cmdCancel_Click
    Else
        'LogAudit "E", "USER MODULES"

        gconDMIS.Execute "UPDATE ALL_RAMS_MODULES set" & _
                       " MODULE_TYPE = " & vtxtModeType & _
                       " where MODULEID = " & labID.Caption

        MessagePop RecSaveInfo, "Module Updated", " Module Information Sucessfully Updated"
    End If

    rsRefresh
    RaiseEvent ChangedRecord(True)

    On Error Resume Next
    rsModules.Find "id =" & labID.Caption
    cmdCancel.Value = True
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11

    Call FillCombo("Select ID, ModuleName from ALL_PROFILE", 0, 1, cboMainModuleName)

    With cboModuleType
        .AddItem ("SYSTEM")
        .AddItem ("DATA ENTRY")
        .AddItem ("TRANSACTION")
        .AddItem ("SEARCH")
        .AddItem ("INQUIRY")
        .AddItem ("REPORTS")
        .AddItem ("PROCESSING")
    End With


    rsRefresh
    Frame1.Enabled = False
    txtSearch.Text = ""
    InitMemVars

    If ModuleID > 0 Then
        Dim TEMPRS                                     As ADODB.Recordset
        Set TEMPRS = rsModules.Clone(adLockReadOnly)
        TEMPRS.Find ("MODULEID=" & ModuleID)
        rsModules.Bookmark = TEMPRS.Bookmark
        Set TEMPRS = Nothing
        cmdEdit_Click
    End If
    lstGrid.ColumnHeaders(1).Width = lstGrid.Width * 0.95
    StoreMemvars
    FillSearchGrid ""
    Screen.MousePointer = 0
End Sub

Sub InitMemVars()
    txtModuleName.Text = ""
    cboMainModuleName.Text = ""
    cboModuleType.ListIndex = -1
End Sub

Sub StoreMemvars()
    If Not rsModules.EOF And Not rsModules.BOF Then
        labID.Caption = rsModules!ModuleID
        txtModuleName.Text = Null2String(rsModules!DESCRIPTIONS)
        cboMainModuleName.Text = Null2String(rsModules!MODULENAME)
        cboModuleType.ListIndex = SelectCombo(cboModuleType, Null2String(rsModules!MODULE_TYPE), False)
    Else
        ShowNoRecord

    End If
End Sub

Sub rsRefresh()
    Set rsModules = New ADODB.Recordset
    rsModules.Open "SELECT A.Modulename,B.MODULEID, B.DESCRIPTIONS,B.MODULE_Type FROM ALL_RAMS_MODULES B INNER JOIN ALL_PROFILE  A ON B.MAINMODULEID=A.ID", gconDMIS, adOpenStatic, adLockReadOnly
End Sub

Private Sub lstGrid_ItemClick(ByVal item As MSComctlLib.ListItem)
    Dim TEMPRS                                         As ADODB.Recordset
    Set TEMPRS = rsModules.Clone(adLockReadOnly)
    TEMPRS.Find ("MODULEID=" & lstGrid.SelectedItem.SubItems(1))
    rsModules.Bookmark = TEMPRS.Bookmark
    Set TEMPRS = Nothing
    StoreMemvars
End Sub
Public Sub EditModule(XModuleID As Long)

    ModuleID = XModuleID

End Sub
Private Sub lstGrid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstGrid
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

Private Sub lstGrid_DblClick()
    cmdEdit.Value = True
End Sub



Private Sub txtsearch_Change()
    FillSearchGrid (txtSearch.Text)
End Sub


Sub FillSearchGrid(xxx As String)
    Dim rsModules                                      As ADODB.Recordset
    lstGrid.Sorted = False: lstGrid.ListItems.Clear
    Set rsModules = New ADODB.Recordset
    Set rsModules = gconDMIS.Execute("SELECT TOP 50 B.DESCRIPTIONS ,  B.MODULEID  FROM ALL_RAMS_MODULES B INNER JOIN ALL_PROFILE  A ON B.MAINMODULEID=A.ID  WHERE B.DESCRIPTIONS like'" & ReplaceQuote(xxx) & "%' order by B.DESCRIPTIONS asc")
    If Not (rsModules.EOF And rsModules.BOF) Then
        Listview_Loadval Me.lstGrid.ListItems, rsModules
        lstGrid.Refresh
    End If
End Sub





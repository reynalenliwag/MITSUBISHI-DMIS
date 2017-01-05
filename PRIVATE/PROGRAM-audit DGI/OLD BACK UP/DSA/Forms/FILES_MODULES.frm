VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmFiles_Modules 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modules"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FILES_MODULES.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDetails 
      Height          =   5205
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   3075
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FILES_MODULES.frx":000C
         Left            =   60
         List            =   "FILES_MODULES.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   270
         Width           =   2925
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   60
         MaxLength       =   35
         TabIndex        =   1
         Top             =   690
         Width           =   2925
      End
      Begin MSComctlLib.ListView lstGrid 
         Height          =   4065
         Left            =   60
         TabIndex        =   2
         Top             =   1080
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   7170
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
         MouseIcon       =   "FILES_MODULES.frx":0010
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
      Height          =   5205
      Left            =   3120
      TabIndex        =   4
      Top             =   -30
      Width           =   5895
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
         Left            =   90
         TabIndex        =   21
         Text            =   "cboModuleType"
         Top             =   4410
         Width           =   3405
      End
      Begin VB.TextBox txtModuleName 
         BackColor       =   &H00FFFFFF&
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
         Left            =   90
         MaxLength       =   185
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Text            =   "FILES_MODULES.frx":0172
         Top             =   2850
         Width           =   3435
      End
      Begin VB.ComboBox cboMainModuleName 
         BackColor       =   &H00FFFFFF&
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
         Left            =   90
         TabIndex        =   19
         Text            =   "cboMainModuleName"
         Top             =   2220
         Width           =   3375
      End
      Begin MSComctlLib.ImageList ImageList 
         Left            =   1590
         Top             =   900
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
               Picture         =   "FILES_MODULES.frx":0178
               Key             =   "AMIS"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FILES_MODULES.frx":2698
               Key             =   "CMIS"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FILES_MODULES.frx":46F5
               Key             =   "CRIS"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FILES_MODULES.frx":6E4C
               Key             =   "CSMS"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FILES_MODULES.frx":8D50
               Key             =   "HRMS"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FILES_MODULES.frx":B0BC
               Key             =   "PMIS"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FILES_MODULES.frx":D32D
               Key             =   "SMIS"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FILES_MODULES.frx":F5D5
               Key             =   "WMIS"
            EndProperty
         EndProperty
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
         Left            =   90
         TabIndex        =   24
         Top             =   4170
         Width           =   2565
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
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   90
         TabIndex        =   23
         Top             =   2610
         Width           =   1155
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
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   90
         TabIndex        =   22
         Top             =   1980
         Width           =   1395
      End
      Begin VB.Image Image1 
         Height          =   1305
         Left            =   120
         Top             =   480
         Width           =   2265
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
         Left            =   270
         TabIndex        =   6
         Top             =   900
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         BorderStyle     =   6  'Inside Solid
         Height          =   1425
         Left            =   90
         Top             =   420
         Width           =   2145
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   -150
      ScaleHeight     =   900
      ScaleWidth      =   7155
      TabIndex        =   7
      Top             =   5190
      Width           =   7155
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   6120
         MouseIcon       =   "FILES_MODULES.frx":11A90
         MousePointer    =   99  'Custom
         Picture         =   "FILES_MODULES.frx":11BE2
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   5430
         MouseIcon       =   "FILES_MODULES.frx":11F48
         MousePointer    =   99  'Custom
         Picture         =   "FILES_MODULES.frx":1209A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   4740
         MouseIcon       =   "FILES_MODULES.frx":123C5
         MousePointer    =   99  'Custom
         Picture         =   "FILES_MODULES.frx":12517
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   4050
         MouseIcon       =   "FILES_MODULES.frx":12873
         MousePointer    =   99  'Custom
         Picture         =   "FILES_MODULES.frx":129C5
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   3360
         MouseIcon       =   "FILES_MODULES.frx":12CD8
         MousePointer    =   99  'Custom
         Picture         =   "FILES_MODULES.frx":12E2A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   2670
         MouseIcon       =   "FILES_MODULES.frx":13124
         MousePointer    =   99  'Custom
         Picture         =   "FILES_MODULES.frx":13276
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   1980
         MouseIcon       =   "FILES_MODULES.frx":135CE
         MousePointer    =   99  'Custom
         Picture         =   "FILES_MODULES.frx":13720
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   5190
      ScaleHeight     =   885
      ScaleWidth      =   2580
      TabIndex        =   14
      Top             =   5160
      Width           =   2580
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   780
         MouseIcon       =   "FILES_MODULES.frx":13A7F
         MousePointer    =   99  'Custom
         Picture         =   "FILES_MODULES.frx":13BD1
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   90
         MouseIcon       =   "FILES_MODULES.frx":13F0F
         MousePointer    =   99  'Custom
         Picture         =   "FILES_MODULES.frx":14061
         Style           =   1  'Graphical
         TabIndex        =   15
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
Attribute VB_Name = "frmFiles_Modules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsModules                           As ADODB.Recordset
Dim AddorEdit                           As String
Dim ModuleID                            As Long
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

Private Sub cboMainModuleName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cmdAdd_Click()
    AddorEdit = "ADD"
    Picture1.Visible = False
    Picture2.Visible = True
    Frame1.Enabled = True
    InitMemVars
    cboMainModuleName.SetFocus

End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    StoreMemVars
End Sub


Private Sub cmdDelete_Click()
    On Error GoTo ErrorCode                                  'AXP063110:17

    If ShowConfirmDelete = True Then
        gconDMIS.Execute ("DELETE FROM ALL_Rams_Modules WHERE MODULEID=" & labID)
        gconDMIS.Execute ("DELETE FROM ALL_Rams_UsersAcess WHERE MODULEID=" & labID)
        rsModules.Requery
        cmdCancel.Value = True
        FillSearchGrid txtSearch
    End If


    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200715:23
Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:

    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
    cboModuleType.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdFind_Click()
    On Error Resume Next                                     'AXP063110:17
    txtSearch.SetFocus
End Sub



'Upating Code       : AXP-0713200715:23
Private Sub cmdNext_Click()
    On Error GoTo ErrorCode:

    rsModules.MoveNext
    If rsModules.EOF Then
        rsModules.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200715:23
Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode:

    rsModules.MovePrevious
    If rsModules.BOF Then
        rsModules.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200715:23
Private Sub cmdSave_Click()
    Dim vtxtModeType                    As String
    Dim vtxtMainModuleID                As Long
    Dim vtxtModuleName                  As String
    On Error GoTo ErrorCode:

    vtxtModeType = N2Str2Null(UCase(cboModuleType.Text))
    vtxtModuleName = N2Str2Null(UCase(txtModuleName.Text))

    If AddorEdit = "ADD" Then
        If cboMainModuleName.ListIndex = -1 Then: cboMainModuleName.SetFocus: Exit Sub
        vtxtMainModuleID = cboMainModuleName.ItemData(cboMainModuleName.ListIndex)
        If vtxtModuleName = "" Then: cboMainModuleName.SetFocus: Exit Sub

        gconDMIS.Execute "INSERT INTO  ALL_RAMS_MODULES (MainModuleNAME,Descriptions,Module_Type)" & _
                       " values('" & LTrim(RTrim(cboMainModuleName)) & "' ," & vtxtModuleName & " ," & vtxtModeType & " )"
        MessagePop RecSaveOk, "New Module Added", " New Module Information Sucessfully Added For " & cboMainModuleName.Text & " System "
        cmdCancel_Click
    Else

        gconDMIS.Execute "UPDATE ALL_RAMS_MODULES SET " & _
                       " DESCRIPTIONS = " & vtxtModuleName & ", " & _
                       " MAINMODULENAME = '" & LTrim(RTrim(cboMainModuleName)) & "', " & _
                       " MODULE_TYPE = " & vtxtModeType & _
                       " WHERE MODULEID = " & labID.Caption

        MessagePop RecSaveInfo, "Module Updated", " Module Information Sucessfully Updated"
    End If


    RaiseEvent ChangedRecord(True)

    On Error Resume Next
    rsModules.Requery
    rsModules.Find "MODULEID=" & labID.Caption
    cmdCancel.Value = True

    Exit Sub
    FillSearchGrid txtSearch
ErrorCode:
    MsgBox Err.Description
    ShowVBError
    Exit Sub

End Sub

Private Sub Combo1_Click()
    FillSearchGrid txtSearch
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me: Exit Sub
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11

    Call FillCombo("Select ID, ModuleName from ALL_PROFILE", 0, 1, cboMainModuleName)
    Combo_Loadval Combo1, gconDMIS.Execute("Select ModuleName from ALL_PROFILE")


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
        Dim TEMPRS                      As ADODB.Recordset
        Set TEMPRS = rsModules.Clone(adLockReadOnly)
        TEMPRS.Find ("MODULEID=" & ModuleID)
        rsModules.Bookmark = TEMPRS.Bookmark
        Set TEMPRS = Nothing
        cmdEdit_Click
    End If


    lstGrid.ColumnHeaders(1).Width = lstGrid.Width * 0.95
    StoreMemVars

    If Combo1.ListCount > 0 Then
        Combo1.ListIndex = 0
    End If

    Screen.MousePointer = 0
End Sub

Sub InitMemVars()
    txtModuleName.Text = ""
    cboMainModuleName.Text = ""
    cboModuleType.ListIndex = -1
End Sub

Sub StoreMemVars()
    If Not rsModules.EOF And Not rsModules.BOF Then
        labID.Caption = rsModules!ModuleID
        txtModuleName.Text = Null2String(rsModules!DESCRIPTIONS)
        cboMainModuleName.Text = Null2String(rsModules!MAINMODULENAME)
        cboModuleType = Null2String(rsModules!MODULE_TYPE)
    Else
        ShowNoRecord

    End If
End Sub

Sub rsRefresh()
    Set rsModules = New ADODB.Recordset
    rsModules.Open "SELECT * FROM ALL_RAMS_MODULES ORDER BY MODULEID DESC ", gconDMIS, adOpenStatic, adLockReadOnly
End Sub

Private Sub lstGrid_ItemClick(ByVal item As MSComctlLib.ListItem)
    Dim TEMPRS                          As ADODB.Recordset
    On Error GoTo ErrorCode                                  'AXP063110:17

    Set TEMPRS = rsModules.Clone(adLockReadOnly)
    TEMPRS.Find ("MODULEID=" & lstGrid.SelectedItem.SubItems(1))
    rsModules.Bookmark = TEMPRS.Bookmark
    Set TEMPRS = Nothing

    StoreMemVars


    Exit Sub
ErrorCode:
    ShowVBError
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



Private Sub txtModuleName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtModuleName_LostFocus()
    txtModuleName = LTrim(RTrim(txtModuleName))
End Sub

Private Sub txtsearch_Change()
    FillSearchGrid (txtSearch.Text)
End Sub


Sub FillSearchGrid(xxx As String)
    Dim rsModules                       As ADODB.Recordset
    lstGrid.Sorted = False: lstGrid.ListItems.Clear
    Set rsModules = New ADODB.Recordset
    If LTrim(RTrim(Combo1.Text)) = "" Then
        Set rsModules = gconDMIS.Execute("SELECT  DESCRIPTIONS ,  MODULEID  FROM ALL_RAMS_MODULES WHERE DESCRIPTIONS like'" & ReplaceQuote(xxx) & "%' order by DESCRIPTIONS asc")
    Else
        Set rsModules = gconDMIS.Execute("SELECT  DESCRIPTIONS ,  MODULEID  FROM ALL_RAMS_MODULES WHERE MAINMODULENAME='" & Repleys(Combo1) & "' AND DESCRIPTIONS like'" & ReplaceQuote(xxx) & "%' order by DESCRIPTIONS asc")
    End If


    If Not (rsModules.EOF And rsModules.BOF) Then
        Listview_Loadval Me.lstGrid.ListItems, rsModules
        lstGrid.Refresh
    End If
End Sub





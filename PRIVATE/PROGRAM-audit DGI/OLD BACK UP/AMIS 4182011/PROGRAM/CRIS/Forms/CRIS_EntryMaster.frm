VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "SHORTC~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCRIS_EntryMaster 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5655
   ClientLeft      =   75
   ClientTop       =   435
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "CRIS_EntryMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5910
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   5715
      Begin VB.TextBox txtDesc 
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
         Left            =   1215
         MaxLength       =   60
         TabIndex        =   2
         Top             =   180
         Width           =   4275
      End
      Begin VB.TextBox txtNotes 
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
         ForeColor       =   &H00701E2A&
         Height          =   780
         Left            =   1215
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   675
         Width           =   4350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
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
         Height          =   210
         Index           =   1
         Left            =   585
         TabIndex        =   4
         Top             =   675
         Width           =   555
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   270
         Width           =   1425
      End
   End
   Begin VB.Frame fraDetails 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   60
      TabIndex        =   6
      Top             =   1935
      Width           =   5715
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
         Height          =   345
         Left            =   135
         MaxLength       =   35
         TabIndex        =   7
         Top             =   225
         Width           =   5535
      End
      Begin MSComctlLib.ListView lvMaster 
         Height          =   1965
         Left            =   90
         TabIndex        =   8
         Top             =   675
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   3466
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "CRIS_EntryMaster.frx":08CA
         NumItems        =   0
      End
   End
   Begin VB.PictureBox PicSave 
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
      Height          =   1005
      Left            =   4320
      ScaleHeight     =   1005
      ScaleWidth      =   1800
      TabIndex        =   9
      Top             =   4620
      Width           =   1800
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
         Left            =   45
         MouseIcon       =   "CRIS_EntryMaster.frx":0A2C
         MousePointer    =   99  'Custom
         Picture         =   "CRIS_EntryMaster.frx":0B7E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   75
         Width           =   705
      End
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
         Left            =   810
         MouseIcon       =   "CRIS_EntryMaster.frx":0ECE
         MousePointer    =   99  'Custom
         Picture         =   "CRIS_EntryMaster.frx":1020
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   75
         Width           =   705
      End
   End
   Begin VB.PictureBox PicAdd 
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
      Left            =   90
      ScaleHeight     =   885
      ScaleWidth      =   6075
      TabIndex        =   12
      Top             =   4665
      Width           =   6075
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
         MouseIcon       =   "CRIS_EntryMaster.frx":135E
         MousePointer    =   99  'Custom
         Picture         =   "CRIS_EntryMaster.frx":14B0
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   45
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
         MouseIcon       =   "CRIS_EntryMaster.frx":180F
         MousePointer    =   99  'Custom
         Picture         =   "CRIS_EntryMaster.frx":1961
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   45
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
         Left            =   1440
         MouseIcon       =   "CRIS_EntryMaster.frx":1CB9
         MousePointer    =   99  'Custom
         Picture         =   "CRIS_EntryMaster.frx":1E0B
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   45
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
         Left            =   2157
         MouseIcon       =   "CRIS_EntryMaster.frx":2105
         MousePointer    =   99  'Custom
         Picture         =   "CRIS_EntryMaster.frx":2257
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   45
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
         Left            =   2866
         MouseIcon       =   "CRIS_EntryMaster.frx":256A
         MousePointer    =   99  'Custom
         Picture         =   "CRIS_EntryMaster.frx":26BC
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   45
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
         Left            =   3555
         MouseIcon       =   "CRIS_EntryMaster.frx":2A18
         MousePointer    =   99  'Custom
         Picture         =   "CRIS_EntryMaster.frx":2B6A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   45
         Width           =   705
      End
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
         Left            =   4995
         MouseIcon       =   "CRIS_EntryMaster.frx":2E95
         MousePointer    =   99  'Custom
         Picture         =   "CRIS_EntryMaster.frx":2FE7
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   45
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
         Left            =   4284
         MouseIcon       =   "CRIS_EntryMaster.frx":334D
         MousePointer    =   99  'Custom
         Picture         =   "CRIS_EntryMaster.frx":349F
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   45
         Width           =   705
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
      Height          =   420
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   10500
      _Version        =   655364
      _ExtentX        =   18521
      _ExtentY        =   741
      _StockProps     =   14
      Caption         =   "asfdafds"
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      ForeColor       =   -2147483630
   End
End
Attribute VB_Name = "frmCRIS_EntryMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMaster                                 As ADODB.Recordset
Private m_eAddorEdit                         As TriadSetting
Private m_MasterType                         As String
Private m_lD                                 As Long
Private m_MasterTypeID                       As Long
Event ChangedData(oRS As ADODB.Recordset)
Private Sub Form_Unload(Cancel As Integer)
    ID = 0
    MasterType = vbNullString
    AddorEdit = DefaultState
End Sub
Public Property Get ID() As Long
    ID = m_lD
End Property
Public Property Let ID(ByVal lID As Long)
    m_lD = lID
End Property
Public Property Get MasterType() As String
    MasterType = m_MasterType
End Property
Public Property Let MasterType(ByVal sMasterType As String)
    m_MasterType = sMasterType
End Property
Public Property Get AddorEdit() As TriadSetting
    AddorEdit = m_eAddorEdit
End Property
Public Property Let AddorEdit(ByVal eAddorEdit As TriadSetting)
    m_eAddorEdit = eAddorEdit
End Property


Private Sub Form_Load()
    rsRefresh
    Frame1.Enabled = False
    ShortcutCaption.Caption = m_MasterType
    InitMemvars
    StoreMemVars
End Sub

Private Sub cmdAdd_Click()
    AddorEdit = AddState
    Frame1.Enabled = True
    picAdd.Visible = False
    picSave.Visible = True
    InitMemvars
    On Error GoTo adder:
    txtDesc.SetFocus
    Exit Sub
adder:
    Err.Clear


End Sub
Private Sub cmdCancel_Click()
    If cmdCancel.Tag = "@EXIT" Then: Unload Me: Exit Sub
    If lvMaster.ListItems.Count = 0 Then
        Unload Me
        Exit Sub
    End If

    Frame1.Enabled = False
    picAdd.Visible = True
    picSave.Visible = False
    StoreMemVars
End Sub
Private Sub cmdDelete_Click()
    '    On Error GoTo ErrorCode
    If lvMaster.SelectedItem Is Nothing Then: Exit Sub

    If Not rsMaster.BOF Or Not rsMaster.EOF Then
        If ShowConfirmDelete = True Then
            oConSQL.Execute "delete from CRIS_MASTERSETUP where [ID] = " & CLng(lvMaster.SelectedItem.ListSubItems(1).Text)
            ShowDeletedMsg
            m_lD = 0
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    rsRefresh
    StoreMemVars
    Exit Sub
    'ErrorCode:
    '    ShowVBError
    '    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    AddorEdit = EditState
    Frame1.Enabled = True
    picAdd.Visible = False
    picSave.Visible = True
End Sub
Private Sub cmdEXIT_Click()
    Unload Me
End Sub
Private Sub cmdFind_CLick()
    txtSearch.SetFocus
End Sub
Private Sub cmdNext_Click()
    rsMaster.MoveNext
    If rsMaster.EOF Then
        rsMaster.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub
Private Sub cmdPrevious_Click()
    rsMaster.MovePrevious
    If rsMaster.BOF Then
        rsMaster.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    Dim Description                          As String
    Dim notes                                As String
    Description = N2Str2Null(StrConv(txtDesc.Text, vbProperCase))
    notes = N2Str2Null(txtNotes.Text)
    If txtDesc.Text = vbNullString Then
        ShowIsRequiredMsg "Color Code and Description"
        txtDesc.SetFocus
        Exit Sub
    End If

    If AddorEdit = AddState Then
        oConSQL.Execute "Insert into CRIS_MasterSetup " & _
                      " (Master_Type_ID,Master_Type_Data, Master_Type_Notes)" & _
                      " values (" _
                      & m_MasterTypeID & ", " _
                      & Description & ", " _
                      & notes & ")"
        MasRecordSet.Requery
    Else
        oConSQL.Execute "update CRIS_MasterSetup set " & _
                      " Master_Type_Data= " & Description & _
                      " , Master_Type_Notes= " & notes & _
                      " where id = " & m_lD

        MasRecordSet.Requery

    End If
    rsRefresh
    cmdCancel.Value = True
    RaiseEvent ChangedData(MasRecordSet)
End Sub


Sub InitMemvars()
    txtDesc.Text = vbNullString
    txtSearch.Text = vbNullString
    txtNotes.Text = vbNullString
End Sub
Sub StoreMemVars()
    If Not rsMaster.EOF And Not rsMaster.BOF Then
        If m_lD > 0 Then
            rsMaster.Filter = "ID=" & m_lD
        End If

        txtDesc.Text = rsMaster![Description]
        txtNotes.Text = Null2String(rsMaster![notes])
        m_lD = CLng(rsMaster![ID])
        m_MasterTypeID = CLng(rsMaster![MasterID])
    Else
        Dim temprs                           As ADODB.Recordset

        Set temprs = GetRS("Select ID from CRIS_MasterType where Master_Type='" & m_MasterType & "'")
        If Not (temprs.EOF And temprs.BOF) Then
            m_MasterTypeID = temprs.Collect(0)
        Else
            MessagePop InfoStop, "Empty Type", "Cannot Add Record. Please Enter Master Set File First", 1500
            Unload Me
            Exit Sub
        End If
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Dim SQL                                  As String
    SQL = " SELECT  " & _
        " MS.ID, " & _
        " MS.Master_Type_ID  as MasterID, " & _
        " MS.Master_Type_Data as [Description], " & _
        " MS.Master_Type_Notes as [Notes] " & _
        " From  " & _
        " CRIS_MasterType MT " & _
        " Inner Join " & _
        " CRIS_MasterSetUP MS " & _
        " ON " & _
        " MT.ID = MS.Master_Type_ID " & _
        " Where MT.Master_Type = '" & m_MasterType & "'"
    Set rsMaster = GetRS(SQL)
    flex_FillListView rsMaster.Clone(adLockReadOnly), lvMaster, True, True
    ConfigHeaders lvMaster, "10,0,0,86,0"
End Sub

Private Sub lvMaster_ItemClick(ByVal Item As MSComctlLib.ListItem)
    m_lD = CLng(Item.ListSubItems(1).Text)
    m_MasterTypeID = CLng(Item.ListSubItems(2).Text)
    txtDesc.Text = Item.ListSubItems(3).Text
    txtNotes.Text = Item.ListSubItems(4).Text
    m_eAddorEdit = EditState
End Sub
Private Sub lvMaster_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvMaster
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
Private Sub lvMaster_DblClick()
    If lvMaster.SelectedItem Is Nothing Then
        Exit Sub
    End If
    cmdEdit.Value = True
End Sub

Private Sub txtSearch_Change()
    Dim TempSQL                              As String
    Dim temprs                               As ADODB.Recordset
    If txtSearch.Text = vbNullString Then
        flex_FillListView rsMaster.Clone(adLockReadOnly), lvMaster, False, False
    Else
        TempSQL = " SELECT  " & _
                " MS.ID, " & _
                " MS.Master_Type_ID  as MasterID, " & _
                " MS.Master_Type_Data as [Description], " & _
                " MS.Master_Type_Notes as [Notes] " & _
                " From  " & _
                " CRIS_MasterType MT " & _
                " Inner Join " & _
                " CRIS_MasterSetUP MS " & _
                " ON " & _
                " MT.ID = MS.Master_Type_ID " & _
                " Where MT.Master_Type ='" & m_MasterType & "'" & _
                " And  MS.Master_Type_Data Like '" & txtSearch.Text & "%'"
        Set temprs = GetRS(TempSQL)
        If Not (temprs.EOF = True And temprs.BOF = True) Then
            flex_FillListView temprs, lvMaster, True
        End If

        ConfigHeaders lvMaster, "10,0,0,86,0"
    End If
End Sub


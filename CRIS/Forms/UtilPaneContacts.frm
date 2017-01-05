VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COD3C4~1.OCX"
Begin VB.Form frmCRIS_PaneContacts 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   8850
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   3030
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddProspect 
      Height          =   315
      Left            =   2280
      Picture         =   "UtilPaneContacts.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Add Prospects"
      Top             =   360
      Width           =   285
   End
   Begin VB.CommandButton cmdOption 
      Height          =   285
      Left            =   2595
      Picture         =   "UtilPaneContacts.frx":01CA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Prospect Options"
      Top             =   375
      Width           =   255
   End
   Begin VB.PictureBox picCustList 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6825
      Left            =   0
      ScaleHeight     =   6825
      ScaleWidth      =   3030
      TabIndex        =   1
      Top             =   690
      Width           =   3030
      Begin VB.Label lblCustDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H00C2FAE2&
         Caption         =   "Miscellenous Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   3615
         Width           =   2835
      End
      Begin VB.Label lblProspectMiscellenous 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004A5A8A&
         Height          =   1890
         Left            =   0
         TabIndex        =   11
         Top             =   3825
         Width           =   2835
      End
      Begin VB.Label lblCustDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H00C2FAE2&
         Caption         =   "Prospect (Profile Summary)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   7
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   2835
      End
      Begin VB.Label lblProspectStatus 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   1500
         Left            =   0
         TabIndex        =   6
         Top             =   1710
         Width           =   2835
      End
      Begin VB.Label lblProspectClassification 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         TabIndex        =   5
         Top             =   1215
         Width           =   2835
      End
      Begin VB.Label lblProspectProfile 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H004A5A8A&
         Height          =   780
         Left            =   0
         TabIndex        =   4
         Top             =   210
         Width           =   2835
      End
      Begin VB.Label lblCustDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H00C2FAE2&
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   3
         Top             =   1500
         Width           =   2835
      End
      Begin VB.Label lblCustDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H00C2FAE2&
         Caption         =   "Prospect Classification"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   1005
         Width           =   2835
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption caps 
      Height          =   315
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   3075
      _Version        =   655364
      _ExtentX        =   5424
      _ExtentY        =   556
      _StockProps     =   14
      Caption         =   "Quick Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.26
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   3
   End
   Begin XtremeShortcutBar.ShortcutCaption captionContacts 
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3060
      _Version        =   655364
      _ExtentX        =   5397
      _ExtentY        =   635
      _StockProps     =   14
      Caption         =   "Prospects"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
   End
End
Attribute VB_Name = "frmCRIS_PaneContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents MasterForm                As frmCRIS_EntryMaster
Attribute MasterForm.VB_VarHelpID = -1

Private Sub cmdAddProspect_Click()
    frmCRIS_EntryProfilePersonal.Show
End Sub

Private Sub cmdOption_Click()
    If frmCRIS_DashBoard.lvGrid.Records.Count = 0 Then
    MessagePop InfoVoid, "No Selection", "Select Your Profile First", 1000, 2
    Exit Sub
    End If
    PopupMenu frmCRIS_DashBoard.mnuContextGrid

End Sub



Private Sub MasterForm_AddedData(oRS As ADODB.Recordset)
    FillTree
    Unload MasterForm
End Sub

Private Sub MasterForm_ChangedData(oRS As ADODB.Recordset)
    FillTree
    Unload MasterForm
End Sub
Sub FillTree()
    tvList.Nodes.Clear
    MasRecordSet.Filter = " MasterType='Contact Type' "
    Dim i                                    As Integer
    'tvList.ImageList = imlIcons
    tvList.Nodes.Add , , , "Create a Group"                   ', 3
    tvList.Nodes.Add , , , "All Prospects"                     ', 1
    While Not MasRecordSet.EOF
        tvList.Nodes.Add , , "T" & MasRecordSet!DataID, MasRecordSet!MasterData    ', 1
        MasRecordSet.MoveNext
    Wend

    MasRecordSet.MoveFirst
End Sub



Private Sub tvList_Click()
    If tvList.SelectedItem.key = vbNullString Then

        frmCRIS_DashBoard.lvGrid.FilterText = vbNullString
        frmCRIS_DashBoard.lvGrid.Populate
        
        Exit Sub
    End If
        If frmCRIS_DashBoard.lvGrid.Records.Count <= 0 Then: Exit Sub
    frmCRIS_DashBoard.lvGrid.FilterText = Right(tvList.SelectedItem.key, Len(tvList.SelectedItem.key) - 1)
    frmCRIS_DashBoard.lvGrid.Columns(4).FooterText = "Current Filter: " & tvList.SelectedItem.Text
    frmCRIS_DashBoard.lvGrid.Populate
    
End Sub

Private Sub tvList_DblClick()
    If tvList.Nodes Is Nothing Then: Exit Sub

    If tvList.SelectedItem.Index = 1 Then
        Set MasterForm = New frmCRIS_EntryMaster
        With MasterForm
            .MasterType = "Contact Type"
            .cmdAdd.Value = True
            .cmdCancel.Tag = "@EXIT"
            .Show
        End With
    ElseIf tvList.SelectedItem.Index = 2 Then
        Exit Sub
    Else
        Set MasterForm = frmCRIS_EntryMaster
        With MasterForm
            .MasterType = "Contact Type"
            .ID = Right(tvList.SelectedItem.key, Len(tvList.SelectedItem.key) - 1)
            .cmdEdit.Value = True
            .cmdCancel.Tag = "@EXIT"
            .Show
        End With
    End If
End Sub

Private Sub tvList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then: tvList_Click
End Sub

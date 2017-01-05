VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSMIS_SearchVehicleInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Vehicle Information"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8235
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00FCFCFC&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SearchTab 
      Height          =   6075
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   10716
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "By &Make"
      TabPicture(0)   =   "SearchVehicleInfo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "By M&odel"
      TabPicture(1)   =   "SearchVehicleInfo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "By &Description"
      TabPicture(2)   =   "SearchVehicleInfo.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "By &Prod No."
      TabPicture(3)   =   "SearchVehicleInfo.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture7"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "By &Ignition Key"
      TabPicture(4)   =   "SearchVehicleInfo.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Picture9"
      Tab(4).ControlCount=   1
      Begin VB.PictureBox Picture9 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74910
         ScaleHeight     =   5595
         ScaleWidth      =   7965
         TabIndex        =   21
         Top             =   90
         Width           =   7965
         Begin VB.PictureBox Picture10 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   24
            Top             =   30
            Width           =   1335
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Keyword:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   60
               TabIndex        =   25
               Top             =   0
               Width           =   1125
            End
         End
         Begin VB.TextBox txtIgnitionKey 
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
            Left            =   1380
            TabIndex        =   23
            Top             =   30
            Width           =   6495
         End
         Begin MSComctlLib.ListView lstIgnitionKey 
            Height          =   5025
            Left            =   30
            TabIndex        =   22
            Top             =   480
            Width           =   7845
            _ExtentX        =   13838
            _ExtentY        =   8864
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15920873
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
            MouseIcon       =   "SearchVehicleInfo.frx":008C
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Ign Key"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Make"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Model"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Description"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Prod No"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "ID"
               Object.Width           =   2
            EndProperty
         End
      End
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74910
         ScaleHeight     =   5595
         ScaleWidth      =   7965
         TabIndex        =   16
         Top             =   90
         Width           =   7965
         Begin MSComctlLib.ListView lstProdNo 
            Height          =   5025
            Left            =   30
            TabIndex        =   20
            Top             =   480
            Width           =   7845
            _ExtentX        =   13838
            _ExtentY        =   8864
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15920873
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
            MouseIcon       =   "SearchVehicleInfo.frx":03A6
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Prod No"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Make"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Model"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Description"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Ign Key"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "ID"
               Object.Width           =   2
            EndProperty
         End
         Begin VB.TextBox txtProdNo 
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
            Left            =   1380
            TabIndex        =   19
            Top             =   30
            Width           =   6495
         End
         Begin VB.PictureBox Picture8 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   17
            Top             =   30
            Width           =   1335
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Keyword:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   60
               TabIndex        =   18
               Top             =   0
               Width           =   1125
            End
         End
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74910
         ScaleHeight     =   5595
         ScaleWidth      =   7965
         TabIndex        =   11
         Top             =   90
         Width           =   7965
         Begin VB.TextBox txtDesc 
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
            Left            =   1380
            TabIndex        =   14
            Top             =   30
            Width           =   6495
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   12
            Top             =   30
            Width           =   1335
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Keyword:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   60
               TabIndex        =   13
               Top             =   0
               Width           =   1125
            End
         End
         Begin MSComctlLib.ListView lstDesc 
            Height          =   5025
            Left            =   30
            TabIndex        =   15
            Top             =   480
            Width           =   7845
            _ExtentX        =   13838
            _ExtentY        =   8864
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15920873
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
            MouseIcon       =   "SearchVehicleInfo.frx":06C0
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Description"
               Object.Width           =   6625
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Model"
               Object.Width           =   5381
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Make"
               Object.Width           =   5381
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Prod No."
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Ign Key"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "ID"
               Object.Width           =   2
            EndProperty
         End
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   -74910
         ScaleHeight     =   5595
         ScaleWidth      =   7965
         TabIndex        =   6
         Top             =   90
         Width           =   7965
         Begin VB.TextBox txtModel 
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
            Left            =   1380
            TabIndex        =   9
            Top             =   30
            Width           =   6495
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   7
            Top             =   30
            Width           =   1335
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Keyword:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   60
               TabIndex        =   8
               Top             =   0
               Width           =   1125
            End
         End
         Begin MSComctlLib.ListView lstModel 
            Height          =   5025
            Left            =   30
            TabIndex        =   10
            Top             =   480
            Width           =   7845
            _ExtentX        =   13838
            _ExtentY        =   8864
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15920873
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
            MouseIcon       =   "SearchVehicleInfo.frx":09DA
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Model"
               Object.Width           =   5381
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   6624
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Make"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Prod No."
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Ign Key"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "ID"
               Object.Width           =   2
            EndProperty
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   5595
         Left            =   90
         ScaleHeight     =   5595
         ScaleWidth      =   7965
         TabIndex        =   1
         Top             =   90
         Width           =   7965
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   30
            ScaleHeight     =   315
            ScaleWidth      =   1275
            TabIndex        =   3
            Top             =   30
            Width           =   1335
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Keyword:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   60
               TabIndex        =   4
               Top             =   0
               Width           =   1125
            End
         End
         Begin VB.TextBox txtMake 
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
            Left            =   1380
            TabIndex        =   2
            Top             =   30
            Width           =   6495
         End
         Begin MSComctlLib.ListView lstMake 
            Height          =   5025
            Left            =   30
            TabIndex        =   5
            Top             =   480
            Width           =   7845
            _ExtentX        =   13838
            _ExtentY        =   8864
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15920873
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
            MouseIcon       =   "SearchVehicleInfo.frx":0CF4
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Make"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Description"
               Object.Width           =   6623
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Model"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Prod No."
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Ign Key"
               Object.Width           =   5380
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "ID"
               Object.Width           =   2
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "frmSMIS_SearchVehicleInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMRRINV                                                          As ADODB.Recordset
Dim Y, k                                                              As Long
Attribute k.VB_VarUserMemId = 1073938433

Private Sub Form_Activate()
    Select Case SEARCH_TAB
        Case 0
            On Error Resume Next
            txtMake.SetFocus

        Case 1
            On Error Resume Next
            txtModel.SetFocus

        Case 2
            On Error Resume Next
            txtDesc.SetFocus

        Case 3
            On Error Resume Next
            txtProdNo.SetFocus
        Case 4
            On Error Resume Next
            txtIgnitionKey.SetFocus
    End Select


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Select Case SEARCH_TAB
            Case 0:
                If Trim(txtMake) <> "" Then
                    On Error Resume Next
                    txtMake.SetFocus
                Else
                    Unload Me
                End If
            Case 1:
                If Trim(txtModel) <> "" Then
                    On Error Resume Next
                    txtModel.SetFocus
                Else
                    Unload Me
                End If

            Case 2:
                If Trim(txtDesc) <> "" Then
                    On Error Resume Next
                    txtDesc.SetFocus
                Else
                    Unload Me
                End If
            Case 3:
                If Trim(txtProdNo) <> "" Then
                    On Error Resume Next
                    txtProdNo.SetFocus
                Else
                    Unload Me
                End If
            Case 4:
                If Trim(txtIgnitionKey) <> "" Then
                    On Error Resume Next
                    txtIgnitionKey.SetFocus
                Else
                    Unload Me
                End If
        End Select
    End If
    If Shift = 2 Then
        Select Case KeyCode
            Case vbKeyO: SearchTab.Tab = 0
            Case vbKeyM: SearchTab.Tab = 1
            Case vbKeyD: SearchTab.Tab = 2
            Case vbKeyP: SearchTab.Tab = 3
            Case vbKeyI: SearchTab.Tab = 4
        End Select
        SEARCH_TAB = SearchTab.Tab: SearchTab_Click (SEARCH_TAB)
    End If
End Sub

Private Sub Form_Load()
    CenterMe Screen, Me, 0
    SearchTab.Tab = SEARCH_TAB

    SearchTab_Click SearchTab.Tab

End Sub

Private Sub lstIgnitionKey_DblClick()
    frmSMIS_Trans_MRR.SearchID (Trim(Me.lstIgnitionKey.SelectedItem.SubItems(5)))
    Unload Me
End Sub

Private Sub lstIgnitionKey_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmSMIS_Trans_MRR.SearchID (Trim(Me.lstIgnitionKey.SelectedItem.SubItems(5)))
        Unload Me
    End If
End Sub

Private Sub lstIgnitionKey_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtIgnitionKey.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub LstModel_DblClick()
    frmSMIS_Trans_MRR.SearchID (Trim(Me.lstModel.SelectedItem.SubItems(5)))
    Unload Me
End Sub

Private Sub LstModel_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtModel.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub LstModel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmSMIS_Trans_MRR.SearchID (Trim(Me.lstModel.SelectedItem.SubItems(5)))
        Unload Me
    End If
End Sub

Private Sub lstProdNo_DblClick()
    frmSMIS_Trans_MRR.SearchID (Trim(Me.lstProdNo.SelectedItem.SubItems(5)))
    Unload Me
End Sub

Private Sub lstProdNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmSMIS_Trans_MRR.SearchID (Trim(Me.lstProdNo.SelectedItem.SubItems(5)))
        Unload Me
    End If
End Sub

Private Sub lstProdNo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtProdNo.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

'Upating Code       : AXP-0707200712:46
Private Sub txtIgnitionKey_Change()
    On Error GoTo ErrorCode:

    If txtIgnitionKey = "" Then
        Me.lstIgnitionKey.Sorted = False: Me.lstIgnitionKey.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select IgnKey, make, descript, model, ProdNo ,id from SMIS_MrrInv order by IgnKey asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstIgnitionKey.ListItems, rsMRRINV
        End If
    Else
        Me.lstIgnitionKey.Sorted = False: Me.lstIgnitionKey.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select IgnKey, make, descript, model, ProdNo, id from SMIS_MrrInv WHERE IgnKey like '" & Trim(Me.txtIgnitionKey) & "%' order by IgnKey asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstIgnitionKey.ListItems, rsMRRINV
        End If
    End If





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub txtIgnitionKey_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtIgnitionKey.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then

        If lstIgnitionKey.ListItems.Count > 0 And lstIgnitionKey.Enabled = True Then: lstIgnitionKey.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtModel_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtModel.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If lstModel.ListItems.Count > 0 And lstModel.Enabled = True Then: lstModel.SetFocus

    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

'Upating Code       : AXP-0707200712:46
'Upating Code       : AXP-0707200712:46
Private Sub txtModel_Change()
    On Error GoTo ErrorCode:

    If txtModel = "" Then
        Me.lstModel.Sorted = False: Me.lstModel.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select model, descript, make, ProdNo, IgnKey ,id from SMIS_MrrInv order by model asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstModel.ListItems, rsMRRINV
        End If
    Else
        Me.lstModel.Sorted = False: Me.lstModel.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select model, descript, make, ProdNo, IgnKey, id from SMIS_MrrInv WHERE model like '" & Trim(ReplaceQuote(Me.txtModel)) & "%' order by model asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstModel.ListItems, rsMRRINV
        End If
    End If





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub LstMake_DblClick()
    If lstMake.SelectedItem Is Nothing Then Exit Sub
    frmSMIS_Trans_MRR.SearchID (Trim(Me.lstMake.SelectedItem.SubItems(5)))
    Unload Me
End Sub

Private Sub LstMake_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtModel.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub LstMake_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmSMIS_Trans_MRR.SearchID (Trim(Me.lstMake.SelectedItem.SubItems(5)))
        Unload Me
    End If
End Sub

Private Sub txtMake_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtMake.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then

        If lstMake.ListItems.Count > 0 And lstMake.Enabled = True Then: lstMake.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

'Upating Code       : AXP-0707200712:46
Private Sub txtMake_Change()
    On Error GoTo ErrorCode:

    If txtMake = "" Then
        Me.lstMake.Sorted = False: Me.lstMake.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select make, descript, model, ProdNo, IgnKey ,id from SMIS_MrrInv order by make asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstMake.ListItems, rsMRRINV
        End If
    Else
        Me.lstMake.Sorted = False: Me.lstMake.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select make,descript,model, ProdNo, IgnKey, id from SMIS_MrrInv WHERE make like '" & Trim(ReplaceQuote(Me.txtMake)) & "%' order by make asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstMake.ListItems, rsMRRINV
        End If
    End If





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub LstDesc_DblClick()
    frmSMIS_Trans_MRR.SearchID (Trim(Me.lstDesc.SelectedItem.SubItems(5)))
    Unload Me
End Sub

Private Sub LstDesc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        On Error Resume Next
        txtDesc.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub LstDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        frmSMIS_Trans_MRR.SearchID (Trim(Me.lstDesc.SelectedItem.SubItems(5)))
        Unload Me
    End If
End Sub

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtDesc.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If lstDesc.ListItems.Count > 0 And lstDesc.Enabled = True Then: lstDesc.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

'Upating Code       : AXP-0707200712:46
Private Sub txtDesc_Change()
    On Error GoTo ErrorCode:

    If txtDesc = "" Then
        Me.lstDesc.Sorted = False: Me.lstDesc.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select descript,model,make, ProdNo, IgnKey, ID from SMIS_MrrInv order by descript asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstDesc.ListItems, rsMRRINV
        End If
    Else
        Me.lstDesc.Sorted = False: Me.lstDesc.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select descript, model, make, ProdNo, IgnKey, ID from SMIS_MrrInv WHERE descript like '" & Trim(Me.txtDesc) & "%' order by descript asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstDesc.ListItems, rsMRRINV
        End If
    End If





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub SearchTab_Click(PreviousTab As Integer)
    SEARCH_TAB = SearchTab.Tab

    DoEvents
    Select Case SEARCH_TAB
        Case 0
            txtMake.Enabled = True: lstMake.Enabled = True
            Me.Caption = "Search Item by Make"
            txtMake_Change
            On Error Resume Next
            txtMake.SetFocus

        Case 1
            txtModel.Enabled = True: lstModel.Enabled = True
            Me.Caption = "Search Item by Vehicle Model"
            txtModel_Change
            On Error Resume Next
            txtModel.SetFocus

        Case 2
            txtDesc.Enabled = True: lstDesc.Enabled = True
            Me.Caption = "Search Item by Description"
            txtDesc_Change
            On Error Resume Next
            txtDesc.SetFocus

        Case 3
            txtProdNo.Enabled = True: lstProdNo.Enabled = True
            Me.Caption = "Search Item by Product Number"
            txtProdNo_Change
            On Error Resume Next
            txtProdNo.SetFocus
        Case 4
            txtIgnitionKey.Enabled = True: lstIgnitionKey.Enabled = True
            Me.Caption = "Search Item by Ignition Key"
            On Error Resume Next
            txtIgnitionKey.SetFocus
    End Select
End Sub

'Upating Code       : AXP-0707200712:46
Private Sub txtProdNo_Change()
    On Error GoTo ErrorCode:

    If txtProdNo = "" Then
        Me.lstProdNo.Sorted = False: Me.lstProdNo.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select ProdNo, make, descript, model, IgnKey ,id from SMIS_MrrInv order by prodno asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstProdNo.ListItems, rsMRRINV
        End If
    Else
        Me.lstProdNo.Sorted = False: Me.lstProdNo.ListItems.Clear
        Set rsMRRINV = New ADODB.Recordset
        Set rsMRRINV = gconDMIS.Execute("select ProdNo, make,descript,make, IgnKey, id from SMIS_MrrInv WHERE prodno like '" & Trim(Me.txtProdNo) & "%' order by prodno asc")
        If Not (rsMRRINV.EOF And rsMRRINV.BOF) Then
            Listview_Loadval Me.lstProdNo.ListItems, rsMRRINV
        End If
    End If





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub txtProdNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trim(txtProdNo.Text) = "" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            KeyCode = 0
        End If
    End If
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If lstProdNo.ListItems.Count > 0 And lstProdNo.Enabled = True Then: lstProdNo.SetFocus
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

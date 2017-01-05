VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCSMSCounter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Counter Master File"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Counter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6270
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   135
      ScaleHeight     =   945
      ScaleWidth      =   6495
      TabIndex        =   12
      Top             =   3915
      Width           =   6495
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
         Left            =   5340
         MouseIcon       =   "Counter.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "Counter.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Exit Window"
         Top             =   60
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
         Left            =   4620
         MouseIcon       =   "Counter.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "Counter.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Print this Record"
         Top             =   60
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
         Left            =   3900
         MouseIcon       =   "Counter.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "Counter.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Delete Selected Record"
         Top             =   60
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
         Left            =   3180
         MouseIcon       =   "Counter.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "Counter.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
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
         Left            =   2460
         MouseIcon       =   "Counter.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "Counter.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Add Record"
         Top             =   60
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
         Left            =   1740
         MouseIcon       =   "Counter.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "Counter.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Find a Record"
         Top             =   60
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
         Left            =   1020
         MouseIcon       =   "Counter.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "Counter.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Move to Next Record"
         Top             =   60
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
         Left            =   300
         MouseIcon       =   "Counter.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "Counter.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
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
      Height          =   1395
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   6105
      Begin VB.TextBox txtDescription 
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
         Left            =   1500
         MaxLength       =   50
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   4155
      End
      Begin VB.TextBox txtNextNumber 
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
         Left            =   1500
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox txtModule 
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
         Left            =   1485
         MaxLength       =   4
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   1065
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
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   630
         Width           =   1545
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Module"
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
         Left            =   180
         TabIndex        =   7
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Next Number"
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
         Left            =   180
         TabIndex        =   4
         Top             =   1020
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2505
      Left            =   60
      TabIndex        =   8
      Top             =   1350
      Width           =   6105
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
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   150
         Width           =   5925
      End
      Begin MSComctlLib.ListView lstCounter 
         Height          =   1875
         Left            =   60
         TabIndex        =   10
         Top             =   540
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   3307
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
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
         MouseIcon       =   "Counter.frx":2D71
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MODULE TYPE"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPTION"
            Object.Width           =   6526
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4770
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   21
      Top             =   3915
      Width           =   1800
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
         MouseIcon       =   "Counter.frx":2ED3
         MousePointer    =   99  'Custom
         Picture         =   "Counter.frx":3025
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Cancel"
         Top             =   60
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
         MouseIcon       =   "Counter.frx":3363
         MousePointer    =   99  'Custom
         Picture         =   "Counter.frx":34B5
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Save Entry"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   270
      TabIndex        =   6
      Top             =   420
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   600
      TabIndex        =   5
      Top             =   270
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmCSMSCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCunter                                           As ADODB.Recordset
Dim AddorEdit                                          As String

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "FILES COUNTER") = False Then Exit Sub
    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    On Error Resume Next
    txtModule.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "FILES COUNTER") = False Then Exit Sub

    On Error GoTo Errorcode
    If Not rsCunter.BOF Or Not rsCunter.EOF Then
        If ShowConfirmDelete = True Then
            gconDMIS.Execute "delete from CSMS_Cunter where id = " & labid.Caption
            ShowDeletedMsg
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    rsRefresh
    StoreMemVars
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "FILES COUNTER") = False Then Exit Sub

    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
    txtModule.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
    'Dim findStr As String
    'findStr = InputSpeechBox("Please Input Module Name ...", txtModule.Text)
    'If findStr <> "" Then
    '   On Error GoTo ErrorCode
    '   rsCunter.Bookmark = rsFind(rsCunter.Clone, "Modul", findStr).Bookmark
    'End If
    'StoreMemvars
    'Exit Sub

    'ErrorCode:
    'If Err.Number = 3021 Then
    '   ShowCantFind findStr
    '   Resume Next
    'End If
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rsCunter.MoveNext
    If rsCunter.EOF Then
        rsCunter.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsCunter.MovePrevious
    If rsCunter.BOF Then
        rsCunter.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "FILES COUNTER") = False Then Exit Sub

End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim VTXTModule                                     As String
    Dim vtxtDescription                                As String
    Dim VTXTNextNumber                                 As Long

    If txtNextNumber.Text = "" Then
        ShowIsRequiredMsg "Next Number"
        On Error Resume Next
        txtNextNumber.SetFocus
        Exit Sub
    End If
    If txtDescription.Text = "" Then
        ShowIsRequiredMsg "Description"
        On Error Resume Next
        txtDescription.SetFocus
        Exit Sub
    End If
    If IsNull(txtModule.Text) = True Then
        ShowIsRequiredMsg "Module Code"
        On Error Resume Next
        txtModule.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Dim rsfindDup                              As ADODB.Recordset
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select modul from CSMS_Cunter where Modul = '" & txtModule.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "Module Code already exist!"
                On Error Resume Next
                txtModule.SetFocus
                Exit Sub
            End If
        End If
    End If

    VTXTModule = N2Str2Null(txtModule.Text)
    vtxtDescription = N2Str2Null(txtDescription.Text)
    VTXTNextNumber = NumericVal(txtNextNumber.Text)

    If AddorEdit = "ADD" Then
        If Not rsCunter.EOF And Not rsCunter.BOF Then
            rsCunter.MoveLast
            labid.Caption = NumericVal(rsCunter!ID) + 1
        End If
        gconDMIS.Execute "Insert into CSMS_Cunter " & _
                         "(Modul,Description,NextNumber,LastUpdate,UserCode)" & _
                       " values (" & VTXTModule & ", " & vtxtDescription & ", " & VTXTNextNumber & ", " & N2Str2Null(LOGDATE) & _
                         ", " & N2Str2Null(LOGCODE) & ")"
    Else
        gconDMIS.Execute "update CSMS_Cunter set" & _
                       " Modul = " & VTXTModule & "," & _
                       " Description = " & vtxtDescription & "," & _
                       " NextNumber = " & VTXTNextNumber & "," & _
                       " LastUpdate = " & N2Str2Null(LOGDATE) & "," & _
                       " UserCode = " & N2Str2Null(LOGCODE) & _
                       " where id = " & labid.Caption
    End If
    rsRefresh
    On Error Resume Next
    rsCunter.Find "id =" & labid.Caption
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
    CenterMe frmMain, Me, 1
    rsRefresh
    Frame1.Enabled = False
    textSearch.Text = "":    'Picture3.ZOrder 0
    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub initMemvars()
    txtModule.Text = ""
    txtNextNumber.Text = ""
    txtDescription.Text = ""
End Sub

Sub StoreMemVars()
On Error GoTo Errorcode
    If Not rsCunter.EOF And Not rsCunter.BOF Then
        labid.Caption = rsCunter!ID
        txtModule.Text = Null2String(rsCunter!modul)
        txtDescription.Text = Null2String(rsCunter!Description)
        txtNextNumber.Text = N2Str2IntZero(rsCunter!nextnumber)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
    Exit Sub
Errorcode:
ShowVBError
End Sub

Sub rsRefresh()
    Set rsCunter = New ADODB.Recordset
    On Error Resume Next 'Temporary
    rsCunter.Open "select * from CSMS_Cunter order by id asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCSMSCounter = Nothing
    UnloadForm Me
End Sub

Private Sub lstCounter_GotFocus()
    rsCunter.Bookmark = rsFind(rsCunter.Clone, "Modul", lstCounter.SelectedItem).Bookmark
    StoreMemVars
End Sub

Private Sub lstCounter_ItemClick(ByVal item As MSComctlLib.ListItem)
    rsCunter.Bookmark = rsFind(rsCunter.Clone, "Modul", lstCounter.SelectedItem).Bookmark
    StoreMemVars
End Sub

Private Sub lstCounter_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstCounter
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

Private Sub lstCounter_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstCounter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (textSearch.Text)
    End If
End Sub

Sub FillGrid()
On Error GoTo Errorcode:
    Dim rsCounter                                      As ADODB.Recordset

    lstCounter.Enabled = False

    lstCounter.Sorted = False: lstCounter.ListItems.Clear
    Set rsCounter = New ADODB.Recordset
    Set rsCounter = gconDMIS.Execute("select Modul, Description, ID from CSMS_Cunter order by Modul asc")
    If Not (rsCounter.EOF And rsCounter.BOF) Then
        Listview_Loadval Me.lstCounter.ListItems, rsCounter
        lstCounter.Refresh
        lstCounter.Enabled = True

    End If

    Exit Sub
Errorcode:
    ShowVBError
End Sub

Sub FillSearchGrid(xxx As String)
    Dim rsCounter                                      As ADODB.Recordset
    lstCounter.Sorted = False: lstCounter.ListItems.Clear
    Set rsCounter = New ADODB.Recordset
    xxx = Repleys(LTrim(RTrim(xxx)))
    Set rsCounter = gconDMIS.Execute("select Modul, Description, ID from CSMS_Cunter where Modul like '" & xxx & "%'")
    If Not (rsCounter.EOF And rsCounter.BOF) Then
        Listview_Loadval Me.lstCounter.ListItems, rsCounter
        lstCounter.Refresh
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstCounter.Enabled = True Then
            lstCounter.SetFocus
        End If
    End If
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

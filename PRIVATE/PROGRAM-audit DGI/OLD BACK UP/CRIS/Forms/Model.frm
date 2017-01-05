VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCSMSModel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Model Data Entry"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Model.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   6210
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   540
      ScaleHeight     =   855
      ScaleWidth      =   7755
      TabIndex        =   12
      Top             =   5130
      Width           =   7755
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
         Left            =   4860
         MouseIcon       =   "Model.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Exit Window"
         Top             =   30
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
         Left            =   4170
         MouseIcon       =   "Model.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Print this Record"
         Top             =   30
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
         Left            =   3480
         MouseIcon       =   "Model.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Delete Selected Record"
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
         Left            =   2790
         MouseIcon       =   "Model.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Left            =   2100
         MouseIcon       =   "Model.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Left            =   1410
         MouseIcon       =   "Model.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Left            =   720
         MouseIcon       =   "Model.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   30
         MouseIcon       =   "Model.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   6105
      Begin VB.ComboBox cboMake 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1230
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   1080
         Width           =   4725
      End
      Begin Crystal.CrystalReport rptSModel 
         Left            =   5490
         Top             =   990
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Vehicle Model Master List"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.TextBox txtDescript 
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
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   1230
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   630
         Width           =   4755
      End
      Begin VB.TextBox txtModel 
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
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   1230
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   180
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Make"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   1140
         Width           =   975
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3465
      Left            =   60
      TabIndex        =   9
      Top             =   1590
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
         Left            =   120
         MaxLength       =   35
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   180
         Width           =   5835
      End
      Begin MSComctlLib.ListView lstModel 
         Height          =   2745
         Left            =   90
         TabIndex        =   11
         Top             =   600
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4842
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
         MouseIcon       =   "Model.frx":2D71
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Model"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPTION"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "MAKE"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4695
      ScaleHeight     =   885
      ScaleWidth      =   2940
      TabIndex        =   21
      Top             =   5130
      Width           =   2940
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
         MouseIcon       =   "Model.frx":2ED3
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":3025
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Left            =   30
         MouseIcon       =   "Model.frx":3363
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":34B5
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   -1920
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   -1590
      TabIndex        =   6
      Top             =   270
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmCSMSModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsS_Model                           As ADODB.Recordset
Dim AddorEdit                           As String

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "MODEL") = False Then Exit Sub
    Screen.MousePointer = 11
    rptSModel.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptSModel.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    LogAudit "G", "VEHICLE MODEL LISTING", txtModel
    PrintSQLReport rptSModel, CSMS_REPORT_PATH & "smodel.rpt", "", CSMS_REPORT_CONNECTION, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "MODEL") = False Then Exit Sub
    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    lstModel.Enabled = False
    textSearch.Enabled = False
    On Error Resume Next
    txtModel.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstModel.Enabled = True
    textSearch.Enabled = True
    fraDetails.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "MODEL") = False Then Exit Sub
    On Error GoTo Errorcode
    If Not rsS_Model.BOF Or Not rsS_Model.EOF Then
        MsgSpeechBox "Delete a Record? Are you Sure?"
        If MsgBoxXP("Are you sure?", "Confirm Delete", XP_YesNo, msg_Question) = True Then
            gconDMIS.Execute "delete from CSMS_MODELS where model = " & labID.Caption
            LogAudit "X", "VEHICLE MODEL", txtModel
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
    If Function_Access(LOGID, "Acess_EDIT", "MODEL") = False Then Exit Sub
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
    On Error Resume Next
    txtModel.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rsS_Model.MoveNext
    If rsS_Model.EOF Then
        rsS_Model.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsS_Model.MovePrevious
    If rsS_Model.BOF Then
        rsS_Model.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    If LTrim(RTrim(txtDescript)) = "" Then
        MsgSpeechBox "Model Description is Required!"
        On Error Resume Next
        txtModel.SetFocus
        Exit Sub
    End If
    If LTrim(RTrim(txtModel.Text)) = True Then
        MsgSpeechBox "Model Code must not be empty"
        On Error Resume Next
        txtModel.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Dim rsfindDup               As ADODB.Recordset
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select model from CSMS_MODELS where Model = '" & txtModel.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "Model Code already exist!"
                On Error Resume Next
                txtModel.SetFocus
                Exit Sub
            End If
        End If
    End If

    Dim VTXTModel As String, VTXTDescript As String, VTXTMake As String
    VTXTModel = N2Str2Null(txtModel.Text)
    VTXTDescript = N2Str2Null(txtDescript.Text)
    VTXTMake = N2Str2Null(cboMake.Text)
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into CSMS_MODELS" & _
                       " (Model,Descript,make,Code)" & _
                       " values (" & VTXTModel & ", " & VTXTDescript & "," & VTXTMake & "," & VTXTModel & ")"
    LogAudit "A", "NEW VEHICLE MODEL", txtModel
    Else
        gconDMIS.Execute "update CSMS_MODELS set" & _
                       " CODE = " & VTXTModel & "," & _
                       " MODEL = " & VTXTModel & "," & _
                       " descript = " & VTXTDescript & "," & _
                       " make = " & VTXTMake & _
                       " where model = '" & labID.Caption & "'"
    LogAudit "E", "VEHICLE MODEL", txtModel
    End If
    rsRefresh
    On Error Resume Next
    rsS_Model.Find "model = '" & labID.Caption & "'"
    cmdCancel.Value = True
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Function SetMakeCode(XXX As String) As String
    Dim rsMake                          As ADODB.Recordset
    Set rsMake = New ADODB.Recordset
    Set rsMake = gconDMIS.Execute("Select * from All_Make where make = '" & XXX & "'")
    If Not rsMake.EOF And Not rsMake.BOF Then
        SetMakeCode = Null2String(rsMake!code)
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    rsRefresh
    Frame1.Enabled = False
    textSearch.Text = "":
    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub initMemvars()
    txtModel.Text = ""
    txtDescript.Text = ""
    FillCboMake
End Sub

Sub FillCboMake()
    Dim rsMake                          As ADODB.Recordset
    Set rsMake = New ADODB.Recordset
    Set rsMake = gconDMIS.Execute("Select upper(make) as make from ALL_Make order by Make asc"): cboMake.Clear
    If Not rsMake.EOF And Not rsMake.BOF Then
        rsMake.MoveFirst
        Do While Not rsMake.EOF
            cboMake.AddItem Null2String(rsMake!Make)
            rsMake.MoveNext
        Loop
    End If
End Sub

Sub StoreMemVars()
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        labID.Caption = Null2String(rsS_Model!Model)
        txtModel.Text = Null2String(rsS_Model!Model)
        txtDescript.Text = Null2String(rsS_Model!descript)
        cboMake.Text = Null2String(rsS_Model!Make)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set rsS_Model = New ADODB.Recordset
    rsS_Model.Open "select * from CSMS_MODELS order by DESCRIPT asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCSMSModel = Nothing
End Sub

Private Sub lstModel_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    rsS_Model.Bookmark = rsFind(rsS_Model.Clone, "MODEL", lstModel.SelectedItem).Bookmark
    StoreMemVars
End Sub

Private Sub lstModel_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstModel
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

Private Sub lstModel_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstModel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then: On Error Resume Next: textSearch.SetFocus
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (textSearch.Text)
    End If
End Sub

Sub FillGrid()
    Dim rsModel                         As ADODB.Recordset
    lstModel.Sorted = False: lstModel.ListItems.Clear: lstModel.Enabled = True
    Set rsModel = New ADODB.Recordset
    Set rsModel = gconDMIS.Execute("select model,DESCRIPT,MAKE from CSMS_MODELS order by model asc")
    If Not (rsModel.EOF And rsModel.BOF) Then
        Listview_Loadval Me.lstModel.ListItems, rsModel
        lstModel.Refresh
    Else
        lstModel.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsModel                         As ADODB.Recordset
    lstModel.Sorted = False: lstModel.ListItems.Clear: lstModel.Enabled = True
    Set rsModel = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsModel = gconDMIS.Execute("select model, DESCRIPT,MAKE from CSMS_MODELS where model like'" & XXX & "%'")
    If Not (rsModel.EOF And rsModel.BOF) Then
        Listview_Loadval Me.lstModel.ListItems, rsModel
        lstModel.Refresh
    Else
        lstModel.Enabled = False
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstModel.ListItems.Count > 0 And lstModel.Enabled = True Then
            lstModel.SetFocus
        End If
    End If
End Sub

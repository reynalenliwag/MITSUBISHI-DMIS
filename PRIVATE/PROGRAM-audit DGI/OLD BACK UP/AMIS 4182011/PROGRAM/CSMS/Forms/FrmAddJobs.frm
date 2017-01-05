VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmCSMSAddJobs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job data entry"
   ClientHeight    =   6660
   ClientLeft      =   180
   ClientTop       =   420
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "FrmAddJobs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin FlexCell.Grid jGrid 
      Height          =   5535
      Left            =   0
      TabIndex        =   19
      Top             =   600
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9763
      BackColor2      =   12648384
      Cols            =   6
      DefaultFontSize =   8.25
      GridColor       =   12632256
      Rows            =   30
   End
   Begin VB.PictureBox picSelectedItem 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4665
      Left            =   3480
      ScaleHeight     =   4635
      ScaleWidth      =   3975
      TabIndex        =   9
      Top             =   1170
      Visible         =   0   'False
      Width           =   4005
      Begin VB.CommandButton cmdpicCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1980
         TabIndex        =   15
         ToolTipText     =   "Cancel"
         Top             =   4080
         Width           =   1455
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Height          =   375
         Left            =   390
         TabIndex        =   14
         ToolTipText     =   "Apply"
         Top             =   4080
         Width           =   1455
      End
      Begin VB.OptionButton optMake 
         Caption         =   "Make"
         Height          =   255
         Left            =   1350
         TabIndex        =   11
         Top             =   150
         Width           =   1215
      End
      Begin VB.TextBox txtSearch 
         Height          =   330
         Left            =   270
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   420
         Width           =   3345
      End
      Begin MSComctlLib.ListView lstModel 
         Height          =   3195
         Left            =   240
         TabIndex        =   13
         Top             =   780
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   5636
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
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
         MouseIcon       =   "FrmAddJobs.frx":01CA
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Model"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Make"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Code"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.OptionButton optModel 
         Caption         =   "Model"
         Height          =   255
         Left            =   270
         TabIndex        =   12
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame frm1 
      Appearance      =   0  'Flat
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
      Height          =   645
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   11805
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   345
         Left            =   6480
         TabIndex        =   18
         ToolTipText     =   "Close"
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   345
         Left            =   5700
         TabIndex        =   17
         ToolTipText     =   "Save Detail"
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show"
         Height          =   345
         Left            =   4710
         TabIndex        =   16
         ToolTipText     =   "Show Details"
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdSelectedModel 
         Caption         =   "&Select Model"
         Height          =   345
         Left            =   10500
         TabIndex        =   8
         ToolTipText     =   "Select Model"
         Top             =   180
         Width           =   1215
      End
      Begin VB.ComboBox cboModel 
         Height          =   345
         Left            =   8190
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   180
         Width           =   1875
      End
      Begin VB.CommandButton cmdModel 
         Caption         =   "..."
         Height          =   345
         Left            =   10140
         TabIndex        =   4
         ToolTipText     =   "View Model"
         Top             =   180
         Width           =   285
      End
      Begin VB.ComboBox cboCategory 
         Height          =   345
         Left            =   1320
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   180
         Width           =   3045
      End
      Begin VB.CommandButton cmdJobCategory 
         Caption         =   "..."
         Height          =   345
         Left            =   4380
         TabIndex        =   1
         ToolTipText     =   "View Job Category"
         Top             =   180
         Width           =   285
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   7590
         TabIndex        =   6
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Job Category"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   210
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "<<  Insert key - Add row  >>             <<  Delete Key - Erase row  >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   7
      Top             =   6210
      Width           =   11205
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "&Add"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuSave 
      Caption         =   "&Save!"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuRefresh 
      Caption         =   "&Refresh"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "&Quit"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmCSMSAddJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMatMas                                           As New ADODB.Recordset
Dim RSUPLOAD                                           As New ADODB.Recordset
Dim rsComon                                            As New ADODB.Recordset
Dim rsSearch                                           As New ADODB.Recordset
Dim cnt                                                As Long
Dim xjCode                                             As String
Dim SaveEdit                                           As String

Private Function GetCategoryCode(XXX As String)
    Dim rsGetCatCode                                   As ADODB.Recordset
    Set rsGetCatCode = New ADODB.Recordset
    Set rsGetCatCode = gconDMIS.Execute("select jCat from CSMS_JobCategory where [desc] ='" & XXX & "'")
    If Not (rsGetCatCode.EOF And rsGetCatCode.BOF) Then
        GetCategoryCode = rsGetCatCode![jcat]
    End If
    Set rsGetCatCode = Nothing
End Function

Sub CreateGrid()
    Set rsComon = New ADODB.Recordset
    If cboModel.Text = "All" Then
        Set rsComon = gconDMIS.Execute("Select jModel,[Desc] from CSMS_JobModel Order by [jModel] Asc")
    Else
        Set rsComon = gconDMIS.Execute("Select jModel,[Desc] from CSMS_JobModel where [desc] = '" & Trim(cboModel.Text) & "'")
    End If

    For cnt = 4 To 60
        jGrid.Column(cnt).Width = 0
    Next cnt

    cnt = 4
    On Error Resume Next
    Do Until rsComon.EOF
        jGrid.Cell(0, cnt).Text = rsComon![Desc]
        jGrid.Column(cnt).Width = 30
        cnt = cnt + 1
        rsComon.MoveNext
    Loop
    Set rsComon = Nothing
End Sub

Sub InitGrid()

    With jGrid
        .Cols = 61: .Rows = 2
        .DisplayFocusRect = False: .AllowUserResizing = True
        .BackColorFixed = &HFFCFB5
        .BackColorFixedSel = &H8000000F
        .BackColorBkg = &HF9EFE3
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Cell(0, 1).Text = "Code"
        .Cell(0, 2).Text = "Code"
        .Cell(0, 3).Text = "Job Description"
        For cnt = 4 To 60
            jGrid.Cell(0, cnt).Text = ""
        Next cnt

        .Column(1).CellType = cellTextBox
        .Column(2).CellType = cellTextBox:
        For cnt = 4 To 60
            jGrid.Cell(0, 3).Text = ""
            jGrid.Column(cnt).CellType = cellTextBox: jGrid.Column(cnt).Alignment = cellCenterCenter
        Next cnt

        .Column(0).Width = 15
        .Column(1).Width = 0: .Column(1).Locked = True
        .Column(2).Width = 60
        .Column(3).Width = 270
        For cnt = 4 To 60
            jGrid.Column(cnt).Width = 0
        Next cnt
        .AllowUserSort = False
        .RowHeight(0) = 25
        .Range(1, 60, .Rows - 1, 60).ForeColor = RGB(0, 0, 128)
    End With
    mnuRefresh_Click
End Sub

Sub LoadCBO()
    Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "select jCat,[desc] from CSMS_JobCategory  order by [jcat] asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        rsMatMas.MoveFirst
        cboCategory.Clear
        Do While Not rsMatMas.EOF
            cboCategory.AddItem Null2String(rsMatMas!Desc)
            rsMatMas.MoveNext
        Loop
    End If

    Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "select jModel,[desc] from CSMS_JobModel order by [jModel] asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        rsMatMas.MoveFirst
        cboModel.Clear
        cboModel.AddItem "All"
        cboModel.Text = "All"
        Do While Not rsMatMas.EOF
            cboModel.AddItem Null2String(rsMatMas!Desc)
            rsMatMas.MoveNext
        Loop
    End If

End Sub

Sub GetNextCode()
    Dim rsCreate                                       As ADODB.Recordset
    Set rsCreate = New ADODB.Recordset
    Set rsCreate = gconDMIS.Execute("Select jCode from CSMS_Jobs_Local Order by jCode desc")
    If Not rsCreate.EOF And Not rsCreate.BOF Then
        xjCode = Format(Val(rsCreate![JCode]) + 1, "000000")
    Else
        xjCode = Format(1, "000000")
    End If
End Sub

Private Sub cboModel_Click()
    Screen.MousePointer = 11
    CreateGrid
    Screen.MousePointer = 0
End Sub

Private Sub cmdApply_Click()
    On Error GoTo Errorcode:

    For cnt = 4 To 60
        jGrid.Column(cnt).Width = 0
    Next cnt
    Dim bevvy, xx                                      As Long
    For bevvy = 1 To Me.lstModel.ListItems.Count
        If lstModel.ListItems(bevvy).Checked = True Then
            For xx = 4 To jGrid.Cols - 1
                If jGrid.Column(xx).Width <= 0 Then
                    If lstModel.ListItems(bevvy).Text = jGrid.Cell(0, xx).Text Then
                        jGrid.Column(xx).Width = 30
                        xx = xx + 1
                    End If
                End If
            Next xx
        End If
    Next bevvy
    picSelectedItem.Visible = False
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdJobCategory_Click()
    frmCSMSJobCategory.Show 1
End Sub

Private Sub cmdModel_Click()
    frmCSMSModel.Show 1
End Sub

Private Sub cmdpicCancel_Click()
    picSelectedItem.Visible = False
End Sub

Private Sub cmdSave_Click()
    If Function_Access(LOGID, "Acess_EDIT", "JOBS") = False Then Exit Sub
    mnuSave_Click
End Sub

Private Sub cmdSelectedModel_Click()
    If cboCategory = "" Then
        MsgBox "Category Name please..."
        Exit Sub
    End If
    If picSelectedItem.Visible = True Then
        picSelectedItem.Visible = False
    Else
        picSelectedItem.Visible = True
    End If

    optModel.Value = True
    txtSearch = ""
End Sub

Private Sub cmdShow_Click()


    mnuEdit_Click
End Sub

Private Sub Form_Activate()
    LoadCBO
End Sub

Private Sub Form_Load()
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitGrid
    LoadCBO
    CreateGrid
End Sub

Private Sub jGrid_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Errorcode:

    If jGrid.ActiveCell.Col <= 1 Then
        If KeyCode = vbKeyDelete Then
            If Function_Access(LOGID, "Acess_DELETE", "JOBS") = False Then Exit Sub
            jGrid.Selection.DeleteByRow
        End If
    End If
    If KeyCode = vbKeyInsert Then
        jGrid.AddItem ""
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub mnuAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "JOBS") = False Then Exit Sub
    mnuAdd.Enabled = False
    mnuSave.Enabled = True
    jGrid.Enabled = True
    mnuEdit.Enabled = False
    SaveEdit = "Add"
End Sub

Private Sub mnuEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "JOBS") = False Then Exit Sub
    If cboCategory.Text = "" Then
        MsgBox "Category name please..."
        Exit Sub
    End If
    cboModel.Text = "All"
    cboModel.Enabled = False

    cboModel.Text = ""
    mnuSave.Enabled = True
    mnuEdit.Enabled = False
    mnuAdd.Enabled = False
    jGrid.Enabled = True
    SaveEdit = "Edit"

    Set rsComon = Nothing
    Set rsComon = New ADODB.Recordset
    Set rsComon = gconDMIS.Execute("Select * from CSMS_Jobs_Local where jCat = '" & GetCategoryCode(cboCategory) & "' Order by ID Asc")
    If Not rsComon.EOF And Not rsComon.BOF Then

        Dim xdisc1                                     As String
        Dim xx                                         As Long
        Dim knt                                        As Double
        cnt = 1
        xdisc1 = rsComon![desc1]
        jGrid.Rows = 1
        jGrid.AddItem rsComon![JCode] & vbTab & _
                      rsComon![OPCODE] & vbTab & _
                      xdisc1
        knt = 0
        Do Until rsComon.EOF
            If rsComon![desc1] = xdisc1 Then
                For xx = 4 To jGrid.Cols - 1
                    If jGrid.Column(xx).Width > 0 Then
                        If rsComon!JCode = jGrid.Cell(0, xx).Text Then
                            jGrid.Cell(cnt, xx).Text = rsComon![std_mhrs]

                        End If
                    End If
                Next xx
            End If
            rsComon.MoveNext
            If Not rsComon.EOF Then
                If rsComon![desc1] <> xdisc1 Then
                    xdisc1 = rsComon![desc1]
                    jGrid.AddItem rsComon![JCode] & vbTab & _
                                  rsComon![OPCODE] & vbTab & _
                                  xdisc1
                    cnt = cnt + 1
                    knt = 0
                End If
            End If
        Loop
    End If
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mnuRefresh_Click()
    Screen.MousePointer = 11
    mnuAdd.Enabled = True: mnuSave.Enabled = False: mnuEdit.Enabled = True:
    cboCategory.Text = "": cboModel.Text = ""
    jGrid.Enabled = False: SaveEdit = ""
    cboModel.Enabled = True
    jGrid.Rows = 2: jGrid.Cell(1, 1).Text = "": jGrid.Cell(1, 2).Text = "": jGrid.Cell(1, 3).Text = ""
    For cnt = 4 To 60
        jGrid.Column(cnt).Width = 0
    Next cnt
    For cnt = 4 To 60
        jGrid.Cell(1, cnt).Text = ""
    Next cnt
    Screen.MousePointer = 0
End Sub

Private Sub mnuSave_Click()
    On Error GoTo Errorcode
    If MsgBox("Save entries..." & vbCrLf & vbCrLf & "Are you sure?", vbYesNo + vbQuestion + vbDefaultButton1, "Message Box") = vbNo Then
        Exit Sub
    End If

    Screen.MousePointer = 11
    gconDMIS.Execute ("Delete from CSMS_Jobs_Local where Jcat = " & N2Str2Null(GetCategoryCode(cboCategory)))
    Dim X                                              As Long
    Dim xDesc1, Xcode, xcodeAmt, xhr, VjCode           As String
    Dim xjCat, xjModel, xjCode2, xopcode               As String
    Dim xFLATRATE, xstd_mhrs                           As Double
    GetNextCode
    For cnt = 1 To jGrid.Rows - 1
        xjCode2 = N2Str2Null(jGrid.Cell(cnt, 1).Text)
        xopcode = N2Str2Null(jGrid.Cell(cnt, 2).Text)
        xDesc1 = N2Str2Null(jGrid.Cell(cnt, 3).Text)
        For X = 4 To jGrid.Cols - 1
            If jGrid.Column(X).Width > 0 Then
                Set rsComon = New ADODB.Recordset
                Set rsComon = gconDMIS.Execute("Select jModel,[Desc] from CSMS_JobModel where [desc] ='" & jGrid.Cell(0, X).Text & "'")
                If Not rsComon.EOF And Not rsComon.BOF Then
                    Xcode = N2Str2Null(rsComon![jmodel])
                    xhr = NumericVal(jGrid.Cell(cnt, X).Text)
                    If jGrid.Cell(cnt, X).Text <> "" Then
                        xjCat = N2Str2Null(GetCategoryCode(cboCategory))
                        xjModel = Xcode
                        xstd_mhrs = xhr


                        gconDMIS.Execute "Insert into CSMS_Jobs_Local " & _
                                       " (opcode,jCat,jModel,jCode,Desc1,std_mhrs)" & _
                                       " values(" & xopcode & "," & xjCat & "," & xjModel & ",'" & xjCode & "'," & xDesc1 & "," & xstd_mhrs & ")"
                        xjCode = Format(Val(xjCode) + 1, "000000")

                    End If
                End If
            End If
        Next X
    Next
    LogAudit "A", "JOB DETAIL ", "FOR THE MODEL " & cboModel
    CmdShow.Value = True
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub txtSEARCH_Change()
    If optModel.Value = True Then
        lstModel.Sorted = False: lstModel.ListItems.Clear
        Set RSUPLOAD = New ADODB.Recordset
        Set RSUPLOAD = gconDMIS.Execute("Select JobModel,[Make],jModel from CSMS_Jobs_Local where jcat = '" & GetCategoryCode(cboCategory) & "' and JobModel like '" & txtSearch & "%' group by JobModel,[Make],jModel Order by [jModel] Asc")
        If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
            Listview_Loadval Me.lstModel.ListItems, RSUPLOAD
        End If
    ElseIf optMake.Value = True Then
        lstModel.Sorted = False: lstModel.ListItems.Clear
        Set RSUPLOAD = New ADODB.Recordset
        Set RSUPLOAD = gconDMIS.Execute("Select JobModel,[Make],jModel from CSMS_Jobs_Local where jcat = '" & GetCategoryCode(cboCategory) & "' and [Make] like '" & txtSearch & "%' group by JobModel,[Make],jModel Order by [jModel] Asc")
        If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
            Listview_Loadval Me.lstModel.ListItems, RSUPLOAD
        End If
    End If
End Sub


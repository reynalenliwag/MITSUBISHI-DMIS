VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRMSCodes_Adjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjustment Codes"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7905
   ForeColor       =   &H00D8E9EC&
   Icon            =   "Codes_Adjustment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   7905
   Visible         =   0   'False
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   2685
      Left            =   2040
      ScaleHeight     =   2685
      ScaleWidth      =   5865
      TabIndex        =   7
      Top             =   1140
      Width           =   5865
      Begin MSComctlLib.ListView lstCodes_Adjustment 
         Height          =   2565
         Left            =   60
         TabIndex        =   8
         Top             =   30
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   4524
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
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
         MouseIcon       =   "Codes_Adjustment.frx":0442
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DEPARTMENT NAME"
            Object.Width           =   6702
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2190
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   10
      Top             =   3810
      Width           =   5580
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
         MouseIcon       =   "Codes_Adjustment.frx":05A4
         MousePointer    =   99  'Custom
         Picture         =   "Codes_Adjustment.frx":06F6
         Style           =   1  'Graphical
         TabIndex        =   18
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
         MouseIcon       =   "Codes_Adjustment.frx":0A5C
         MousePointer    =   99  'Custom
         Picture         =   "Codes_Adjustment.frx":0BAE
         Style           =   1  'Graphical
         TabIndex        =   17
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
         MouseIcon       =   "Codes_Adjustment.frx":0F14
         MousePointer    =   99  'Custom
         Picture         =   "Codes_Adjustment.frx":1066
         Style           =   1  'Graphical
         TabIndex        =   16
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
         MouseIcon       =   "Codes_Adjustment.frx":1391
         MousePointer    =   99  'Custom
         Picture         =   "Codes_Adjustment.frx":14E3
         Style           =   1  'Graphical
         TabIndex        =   15
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
         MouseIcon       =   "Codes_Adjustment.frx":183F
         MousePointer    =   99  'Custom
         Picture         =   "Codes_Adjustment.frx":1991
         Style           =   1  'Graphical
         TabIndex        =   14
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
         MouseIcon       =   "Codes_Adjustment.frx":1CA4
         MousePointer    =   99  'Custom
         Picture         =   "Codes_Adjustment.frx":1DF6
         Style           =   1  'Graphical
         TabIndex        =   13
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
         MouseIcon       =   "Codes_Adjustment.frx":20F0
         MousePointer    =   99  'Custom
         Picture         =   "Codes_Adjustment.frx":2242
         Style           =   1  'Graphical
         TabIndex        =   12
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
         MouseIcon       =   "Codes_Adjustment.frx":259A
         MousePointer    =   99  'Custom
         Picture         =   "Codes_Adjustment.frx":26EC
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   4470
      Left            =   60
      ScaleHeight     =   4410
      ScaleWidth      =   1845
      TabIndex        =   9
      Top             =   180
      Width           =   1905
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   6960
         Left            =   0
         Picture         =   "Codes_Adjustment.frx":2A4B
         Top             =   0
         Width           =   9915
      End
   End
   Begin VB.PictureBox picCodes_Adjustment 
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   2040
      ScaleHeight     =   1035
      ScaleWidth      =   5865
      TabIndex        =   2
      Top             =   120
      Width           =   5865
      Begin VB.TextBox txtCodes 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   0
         Top             =   90
         Width           =   1035
      End
      Begin VB.TextBox txtDescription 
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
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   1
         Top             =   510
         Width           =   4425
      End
      Begin Crystal.CrystalReport rptAdjustment 
         Left            =   5310
         Top             =   30
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Height          =   255
         Index           =   0
         Left            =   -60
         TabIndex        =   6
         Top             =   120
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Height          =   255
         Left            =   30
         TabIndex        =   5
         Top             =   570
         Width           =   1155
      End
      Begin VB.Label labID 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   4
         Top             =   570
         Width           =   225
      End
      Begin VB.Label labIDprev 
         Caption         =   "IDprev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3780
         TabIndex        =   3
         Top             =   570
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   6345
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   19
      Top             =   3825
      Width           =   1440
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
         MouseIcon       =   "Codes_Adjustment.frx":167A8
         MousePointer    =   99  'Custom
         Picture         =   "Codes_Adjustment.frx":168FA
         Style           =   1  'Graphical
         TabIndex        =   21
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
         MouseIcon       =   "Codes_Adjustment.frx":16C38
         MousePointer    =   99  'Custom
         Picture         =   "Codes_Adjustment.frx":16D8A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmHRMSCodes_Adjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCodes_Adjustment                                                As ADODB.Recordset
Dim AddorEdit                                                         As String
Dim ADD_EDIT_GROUP                                                    As String
'last update 3/16/07       --    Jonathan               ---     20070316

Function CheckIfCodeAdjustmentAlreadyExist() As Boolean
    Dim RSTMP                                                         As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select Codes From HRMS_Codes_Adjustment Where Codes = '" & txtCodes.Text & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        CheckIfCodeAdjustmentAlreadyExist = True
    Else
        CheckIfCodeAdjustmentAlreadyExist = False
    End If

    Set RSTMP = Nothing
End Function

Sub GenerateNewAdjustCode()
    Dim RSTMP                                                         As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select Codes From HRMS_Codes_Adjustment Order By Codes Desc")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        txtCodes.Text = Format(Val(RSTMP!codes) + 1, "000")
    Else
        txtCodes.Text = "001"
    End If
    Set RSTMP = Nothing
End Sub

Sub DisableMain(COND As Boolean)
    Picture4.Enabled = COND
    Picture2.Enabled = COND
    Picture1.Enabled = COND
    lstCodes_Adjustment.Enabled = COND

    'PicGroup.Visible = Not Cond
End Sub

Sub DisablePic(COND As Boolean)

End Sub

Sub FillCboGroup()

End Sub

Sub rsRefresh()
    Set rsCodes_Adjustment = New ADODB.Recordset
    rsCodes_Adjustment.Open "select * from HRMS_Codes_Adjustment order by Codes", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub InitMemvars()
    picCodes_Adjustment.Enabled = True
    txtCodes.Text = ""
    txtDescription.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsCodes_Adjustment.EOF And Not rsCodes_Adjustment.BOF Then
        picCodes_Adjustment.Enabled = False
        labID.Caption = rsCodes_Adjustment!ID
        txtCodes.Text = Null2String(rsCodes_Adjustment!codes)
        txtDescription.Text = Null2String(rsCodes_Adjustment!Description)
    Else
        picCodes_Adjustment.Enabled = False
        ShowNoRecord
        'last update 3/16/07       --    Jonathan               ---     20070316
        'bug       ---    continues pop-up of msgbox
        'If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdADD.Value = True Else Unload Me
    End If
End Sub

Sub FillGrid()
    Dim rsCodes_Adjustment2                                           As ADODB.Recordset

    lstCodes_Adjustment.Enabled = False
    lstCodes_Adjustment.Sorted = False: lstCodes_Adjustment.ListItems.Clear
    Set rsCodes_Adjustment2 = New ADODB.Recordset
    Set rsCodes_Adjustment2 = gconDMIS.Execute("select Codes,Description,ID from HRMS_Codes_Adjustment")
    If Not (rsCodes_Adjustment2.EOF And rsCodes_Adjustment2.BOF) Then
        Listview_Loadval Me.lstCodes_Adjustment.ListItems, rsCodes_Adjustment2
        lstCodes_Adjustment.Refresh
        lstCodes_Adjustment.Enabled = True
    End If
End Sub

Private Sub cboGroup_Change()
    'If Not cboGroup.ListCount = 0 Then txtGroup_Codes.Text = Right(cboGroup, 2)
End Sub

Private Sub cboGroup_Click()
    'If Not cboGroup.ListCount = 0 Then txtGroup_Codes.Text = Right(cboGroup, 2)
End Sub

Private Sub cboGroup_LostFocus()
    'If Not cboGroup.ListCount = 0 Then txtGroup_Codes.Text = Right(cboGroup, 2)
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Add", "FILES ADJUSTMENTS") = False Then Exit Sub

    AddorEdit = "ADD"
    InitMemvars
    GenerateNewAdjustCode

    lstCodes_Adjustment.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True

    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub cmdAGroup_Click()
    DisableMain False
End Sub

Private Sub cmdCancel_Click()
    picCodes_Adjustment.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstCodes_Adjustment.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Delete", "FILES ADJUSTMENTS") = False Then Exit Sub

    If Not txtCodes.Text = "" Then
        If ShowConfirmDelete = True Then
            SQL_STATEMENT = "delete from HRMS_Codes_Adjustment where Codes = '" & txtCodes.Text & "'"
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "X", "FILES ADJUSTMENTS", SQL_STATEMENT, txtCodes.Text, "", "", "", ""
            SQL_STATEMENT = ""
            txtCodes.Text = ""
            txtDescription.Text = ""
            ShowDeletedMsg
        End If

        rsRefresh
        StoreMemVars
        FillGrid
    End If
    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Edit", "FILES ADJUSTMENTS") = False Then Exit Sub

    AddorEdit = "EDIT"
    picCodes_Adjustment.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    lstCodes_Adjustment.Enabled = False

    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    UnloadForm Me
End Sub

Private Sub cmdFind_Click()
    MsgBox "Pls use the List view to find...", vbInformation, "Find"
End Sub

Private Sub cmdgAdd_Click()
    ADD_EDIT_GROUP = "ADD"
    DisablePic False
End Sub

Private Sub cmdgCancel_Click()
    DisablePic True
End Sub

Private Sub cmdgExit_Click()
    DisableMain True
End Sub

Private Sub cmdNext_Click()
    rsCodes_Adjustment.MoveNext
    If rsCodes_Adjustment.EOF Then
        rsCodes_Adjustment.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsCodes_Adjustment.MovePrevious
    If rsCodes_Adjustment.BOF Then
        rsCodes_Adjustment.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Print", "FILES ADJUSTMENTS") = False Then Exit Sub


    Screen.MousePointer = 11
    rptAdjustment.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
    rptAdjustment.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
    rptAdjustment.Formulas(2) = "COMPANYTIN = '" & COMPANY_TIN & "'"
    rptAdjustment.Formulas(3) = "PrintedBy = '" & LOGNAME & "'"

    PrintSQLReport rptAdjustment, HRMS_REPORT_PATH & "Adjustment List.rpt", "", DMIS_REPORT_Connection, 1
    LogAudit "V", "PRINT ADJUSTMENT RECORD", ""
    Screen.MousePointer = 0

    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim EXIST                                                         As Boolean

    If AddorEdit = "ADD" Then
        If CheckIfCodeAdjustmentAlreadyExist = True Then
            MsgBox "Adustment Code Already Exist", vbInformation, "Duplicate of Data"
            txtCodes.SetFocus
            Exit Sub
        End If
    End If

    txtCodes.Text = N2Str2Null(txtCodes.Text)
    txtDescription.Text = N2Str2Null(txtDescription.Text)

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "Insert into HRMS_Codes_Adjustment " & _
                        "(Codes,Description) " & _
                      " values (" & txtCodes.Text & ", " & txtDescription.Text & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", "FILES ADJUSTMENTS", SQL_STATEMENT, txtCodes.Text, "", "", "", ""
        SQL_STATEMENT = ""
        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update HRMS_Codes_Adjustment set Description = " & txtDescription.Text & _
                      " where Codes = " & txtCodes & ""
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", "FILES ADJUSTMENTS", SQL_STATEMENT, txtCodes.Text, "", "", "", ""
        SQL_STATEMENT = ""
        ShowSuccessFullyUpdated
    End If

    rsRefresh
    FillGrid
    cmdCancel.Value = True
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (FILES ADJUSTMENTS)"
            Call frmALL_AuditInquiry.DisplayHistory(labID, "FILES ADJUSTMENTS")
    End Select
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    rsRefresh
    StoreMemVars
    FillGrid

    FillCboGroup
    'DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub lstCodes_Adjustment_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstCodes_Adjustment
        .Sorted = True
        If .SortKey = ColumnHeader.INDEX - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.INDEX - 1
        End If
    End With
End Sub

Private Sub lstCodes_Adjustment_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstCodes_Adjustment_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    On Error Resume Next
    Dim INDEX                                                         As Double

    If Not lstCodes_Adjustment.ListItems.count = 0 Then
        With lstCodes_Adjustment
            INDEX = .SelectedItem.INDEX
            txtCodes.Text = .ListItems(INDEX).Text
            txtDescription.Text = .ListItems(INDEX).SubItems(1)
        End With
    End If
    'rsCodes_Adjustment.Bookmark = rsFind(rsCodes_Adjustment.Clone, "Codes", Me.lstCodes_Adjustment(Me.lstCodes_Adjustment.SelectedItem.INDEX).Text).Bookmark
    'Call StoreMemVars
End Sub

Private Sub txtCodes_LostFocus()
    txtCodes.Text = UCase(txtCodes.Text)
End Sub


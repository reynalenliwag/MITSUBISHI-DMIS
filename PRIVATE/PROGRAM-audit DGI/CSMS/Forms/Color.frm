VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSMIS_Files_Color 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Color"
   ClientHeight    =   5370
   ClientLeft      =   75
   ClientTop       =   495
   ClientWidth     =   5835
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Color.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   5835
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   6075
      TabIndex        =   13
      Top             =   4410
      Width           =   6075
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
         Left            =   5010
         MouseIcon       =   "Color.frx":20D2
         MousePointer    =   99  'Custom
         Picture         =   "Color.frx":2224
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Exit Window"
         Top             =   60
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
         Left            =   4320
         MouseIcon       =   "Color.frx":258A
         MousePointer    =   99  'Custom
         Picture         =   "Color.frx":26DC
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Print this Record"
         Top             =   60
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
         Left            =   3630
         MouseIcon       =   "Color.frx":2A42
         MousePointer    =   99  'Custom
         Picture         =   "Color.frx":2B94
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Delete Selected Record"
         Top             =   60
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
         Left            =   2940
         MouseIcon       =   "Color.frx":2EBF
         MousePointer    =   99  'Custom
         Picture         =   "Color.frx":3011
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
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
         Left            =   2250
         MouseIcon       =   "Color.frx":336D
         MousePointer    =   99  'Custom
         Picture         =   "Color.frx":34BF
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Add Record"
         Top             =   60
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
         Left            =   1560
         MouseIcon       =   "Color.frx":37D2
         MousePointer    =   99  'Custom
         Picture         =   "Color.frx":3924
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Find a Record"
         Top             =   60
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
         Left            =   870
         MouseIcon       =   "Color.frx":3C1E
         MousePointer    =   99  'Custom
         Picture         =   "Color.frx":3D70
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Move to Next Record"
         Top             =   60
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
         Left            =   180
         MouseIcon       =   "Color.frx":40C8
         MousePointer    =   99  'Custom
         Picture         =   "Color.frx":421A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   30
      TabIndex        =   2
      Top             =   -60
      Width           =   5715
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   4350
         Top             =   270
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox txtColor_code 
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
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   180
         Width           =   1200
      End
      Begin VB.TextBox txtColor_desc 
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
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   4440
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Height          =   285
         Left            =   630
         TabIndex        =   4
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Height          =   285
         Left            =   90
         TabIndex        =   3
         Top             =   660
         Width           =   1425
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3375
      Left            =   30
      TabIndex        =   7
      Top             =   990
      Width           =   5715
      Begin VB.OptionButton optCode 
         Caption         =   "&Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3030
         TabIndex        =   12
         Top             =   180
         Width           =   1245
      End
      Begin VB.OptionButton optDesc 
         Caption         =   "&Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Top             =   210
         Value           =   -1  'True
         Width           =   1305
      End
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
         Left            =   90
         MaxLength       =   35
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   570
         Width           =   5535
      End
      Begin MSComctlLib.ListView lstColor 
         Height          =   2325
         Left            =   60
         TabIndex        =   9
         Top             =   960
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   4101
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Color.frx":4579
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPTION"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Search by:"
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
         Left            =   150
         TabIndex        =   10
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4260
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   21
      Top             =   4455
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
         Left            =   750
         MouseIcon       =   "Color.frx":46DB
         MousePointer    =   99  'Custom
         Picture         =   "Color.frx":482D
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   60
         MouseIcon       =   "Color.frx":4B6B
         MousePointer    =   99  'Custom
         Picture         =   "Color.frx":4CBD
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label labPrev 
      Caption         =   "Label4"
      Height          =   315
      Left            =   8160
      TabIndex        =   6
      Top             =   570
      Width           =   195
   End
   Begin VB.Label labid 
      Caption         =   "Label4"
      Height          =   255
      Left            =   8160
      TabIndex        =   5
      Top             =   690
      Width           =   225
   End
End
Attribute VB_Name = "frmSMIS_Files_Color"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsColor                                            As ADODB.Recordset
Dim ADDOREDIT                                          As String

Private Sub cmdAdd_Click()
    On Error GoTo ERRORCODE:

    If Function_Access(LOGID, "Acess_Add", "VEHICLE COLOR") = False Then Exit Sub

    ADDOREDIT = "ADD"
    Frame1.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    initMemvars
    lstColor.Enabled = False
    txtSearch.Enabled = False
    optDesc.Enabled = False
    optCode.Enabled = False
    On Error Resume Next
    txtColor_code.SetFocus

    Exit Sub
ERRORCODE:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    lstColor.Enabled = True
    txtSearch.Enabled = True
    fraDetails.Enabled = True

    optDesc.Enabled = True
    optCode.Enabled = True

    StoreMemvars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "VEHICLE COLOR") = False Then Exit Sub
    On Error GoTo ERRORCODE
    ''
    If Not rsColor.BOF Or Not rsColor.EOF Then
        If ShowConfirmDelete = True Then
            SQL_STATEMENT = "delete from ALL_Color where id = " & labid.Caption
            gconDMIS.Execute SQL_STATEMENT
            'LogAudit "X", "COLOR MASTER FILE ", txtColor_code
            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("X", "VEHICLE COLOR", SQL_STATEMENT, labid, "", "CODE: " & txtColor_code, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
            ShowDeletedMsg
            FillSearchGrid ""
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    rsRefresh
    StoreMemvars
    Exit Sub

ERRORCODE:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo ERRORCODE:
    If Function_Access(LOGID, "Acess_EDIT", "VEHICLE COLOR") = False Then Exit Sub
    ADDOREDIT = "EDIT"
    Frame1.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    fraDetails.Enabled = False

    Exit Sub
ERRORCODE:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next

    txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsColor.MoveNext
    If rsColor.EOF Then
        rsColor.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
    rsColor.MovePrevious
    If rsColor.BOF Then
        rsColor.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "VEHICLE COLOR") = False Then Exit Sub
    PrintSQLReport CrystalReport1, SMIS_REPORT_PATH & "Listing/Colors.rpt", "", DMIS_REPORT_Connection, 1

    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("V", "VEHICLE COLOR", "", labid, "", "CODE: " & txtColor_code, "", "")
    'NEW LOG AUDIT-----------------------------------------------------
    'LogAudit "V", "COLOR MASTER FILE ", Now()
End Sub

Private Sub cmdSave_Click()
    Dim lng                                            As Integer
    On Error GoTo ERRORCODE:

    If txtColor_code.Text = "" Or txtColor_desc.Text = "" Then
        ShowIsRequiredMsg "Color Code and Description"
        On Error Resume Next
        txtColor_code.SetFocus
        Exit Sub
    End If
    ''''''
    lng = gconDMIS.Execute("select Count(*) from ALL_Color WHERE color_code=" & N2Str2Null(txtColor_code)).Fields(0).Value
    If ADDOREDIT = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exist"
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(Null2String(rsColor!Color_code)) <> UCase(txtColor_code) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exist"
            Exit Sub
        End If
    End If

    If ADDOREDIT = "ADD" Then
        gconDMIS.Execute "Insert into ALL_Color" & _
                       " (color_code,color_desc)" & _
                       " values (" & N2Str2Null(txtColor_code.Text) & ", " & N2Str2Null(txtColor_desc.Text) & ")"
        'LogAudit "A", "COLOR MASTER FILE ", txtColor_code & " " & txtColor_desc
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("A", "VEHICLE COLOR", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtColor_code), "COLOR_CODE", "ALL_COLOR"), "", "CODE: " & txtColor_code, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update ALL_Color set" & _
                      " color_code = " & N2Str2Null(txtColor_code.Text) & "," & _
                      " color_desc = " & N2Str2Null(txtColor_desc.Text) & _
                      " where id = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT
        'LogAudit "E", "COLOR MASTER FILE ", txtColor_code & " " & txtColor_desc

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "VEHICLE COLOR", SQL_STATEMENT, labid, "", "CODE: " & txtColor_code, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        ShowSuccessFullyUpdated
    End If

    rsRefresh
    If ADDOREDIT = "EDIT" Then
        rsColor.Find ("ID=" & labid)
    End If
    cmdCancel.Value = True
    FillSearchGrid ""

    Exit Sub
ERRORCODE:
    ShowVBError
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsColor2                                       As ADODB.Recordset
    lstColor.Sorted = False
    lstColor.ListItems.Clear
    lstColor.Enabled = False
    Set rsColor2 = New ADODB.Recordset

    If optCode.Value = True Then
        Set rsColor2 = gconDMIS.Execute("select  Color_code , color_desc, ID from ALL_Color where Color_code like'" & ReplaceQuote(XXX) & "%' order by color_desc asc")
    Else
        Set rsColor2 = gconDMIS.Execute("select  Color_code , color_desc, ID from ALL_Color where color_desc like'" & ReplaceQuote(XXX) & "%' order by color_desc asc")
    End If

    If Not (rsColor2.EOF And rsColor2.BOF) Then
        Listview_Loadval Me.lstColor.ListItems, rsColor2
        lstColor.Refresh
        lstColor.Enabled = True
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If picAdds.Visible = True And KeyCode = vbKeyEscape Then
        Unload Me
    Else
        MoveKeyPress KeyCode
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (VEHICLE COLOR MASTER FILE)"
            Call frmALL_AuditInquiry.DisplayHistory(labid, "VEHICLE COLOR", "")

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh

    txtSearch.Text = vbNullString
    Frame1.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    initMemvars
    StoreMemvars
    Screen.MousePointer = 0
End Sub

Sub initMemvars()
    txtColor_code.Text = vbNullString
    txtColor_desc.Text = vbNullString
End Sub

Private Sub lstColor_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstColor
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

Private Sub lstColor_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstColor_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsColor.MoveFirst
    rsColor.Find ("ID=" & Item.ListSubItems(2).Text)
    StoreMemvars
End Sub

Private Sub optCode_Click()
    If txtSearch = "" Then FillSearchGrid (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub optDesc_Click()
    If txtSearch = "" Then FillSearchGrid (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Sub rsRefresh()
    Set rsColor = New ADODB.Recordset
    rsColor.Open "select * from ALL_Color order by id DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemvars()
    If Not rsColor.EOF And Not rsColor.BOF Then
        labid.Caption = rsColor!ID
        txtColor_code.Text = Null2String(rsColor!Color_code)
        txtColor_desc.Text = Null2String(rsColor!color_desc)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Private Sub txtColor_code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtColor_desc_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtsearch_Change()
    FillSearchGrid (txtSearch.Text)
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCRIS_ClassifyLeads 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Lead Classification"
   ClientHeight    =   6630
   ClientLeft      =   315
   ClientTop       =   525
   ClientWidth     =   5775
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00F5F5F5&
   Icon            =   "ClassifyLeads.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   5775
   Begin VB.Frame fraDetails 
      Height          =   3375
      Left            =   15
      TabIndex        =   11
      Top             =   2280
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
         Left            =   90
         MaxLength       =   35
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   570
         Width           =   5535
      End
      Begin VB.OptionButton optDesc 
         Caption         =   "&Description"
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Top             =   210
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton optCode 
         Caption         =   "&Code"
         Height          =   375
         Left            =   3030
         TabIndex        =   12
         Top             =   180
         Width           =   1245
      End
      Begin MSComctlLib.ListView lstColor 
         Height          =   2325
         Left            =   60
         TabIndex        =   15
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
         MouseIcon       =   "ClassifyLeads.frx":08CA
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Class"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Search by:"
         Height          =   345
         Left            =   150
         TabIndex        =   16
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   5775
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.TextBox txtCode 
         Height          =   390
         Left            =   2250
         MaxLength       =   10
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   90
         Width           =   3165
      End
      Begin VB.TextBox txtDescription 
         Height          =   390
         Left            =   2250
         MaxLength       =   30
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   540
         Width           =   3165
      End
      Begin VB.Frame Frame1 
         Height          =   1380
         Left            =   45
         TabIndex        =   1
         Top             =   900
         Width           =   5655
         Begin VB.CheckBox chkSO 
            Caption         =   "Sales Order"
            Height          =   225
            Left            =   150
            TabIndex        =   8
            Top             =   225
            Width           =   2190
         End
         Begin VB.CheckBox chkLO 
            Caption         =   "Loan Application"
            Height          =   225
            Left            =   150
            TabIndex        =   7
            Top             =   480
            Width           =   2190
         End
         Begin VB.CheckBox chkTD 
            Caption         =   "Test Drive"
            Height          =   225
            Left            =   150
            TabIndex        =   6
            Top             =   750
            Width           =   2190
         End
         Begin VB.CheckBox chkQU 
            Caption         =   "Quotations"
            Height          =   225
            Left            =   2715
            TabIndex        =   5
            Top             =   165
            Width           =   2190
         End
         Begin VB.CheckBox chkLetter 
            Caption         =   "Letters"
            Height          =   225
            Left            =   2715
            TabIndex        =   4
            Top             =   435
            Width           =   2190
         End
         Begin VB.CheckBox chkEmail 
            Caption         =   "Email"
            Height          =   225
            Left            =   2715
            TabIndex        =   3
            Top             =   690
            Width           =   2190
         End
         Begin VB.CheckBox chkApp 
            Caption         =   "Appointment"
            Height          =   225
            Left            =   150
            TabIndex        =   2
            Top             =   1020
            Width           =   2190
         End
      End
      Begin VB.Label Label2 
         Caption         =   "CLASSIFICATION"
         Height          =   240
         Left            =   150
         TabIndex        =   30
         Top             =   585
         Width           =   1830
      End
      Begin VB.Label Label1 
         Caption         =   "CODE"
         Height          =   240
         Left            =   150
         TabIndex        =   29
         Top             =   150
         Width           =   1830
      End
      Begin VB.Label labid 
         Caption         =   "0"
         Height          =   540
         Left            =   5025
         TabIndex        =   28
         Top             =   1950
         Visible         =   0   'False
         Width           =   315
      End
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   -30
      ScaleHeight     =   945
      ScaleWidth      =   6075
      TabIndex        =   17
      Top             =   5670
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
         MouseIcon       =   "ClassifyLeads.frx":0A2C
         MousePointer    =   99  'Custom
         Picture         =   "ClassifyLeads.frx":0B7E
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Exit Window"
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
         Left            =   4320
         MouseIcon       =   "ClassifyLeads.frx":0EE4
         MousePointer    =   99  'Custom
         Picture         =   "ClassifyLeads.frx":1036
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Left            =   3630
         MouseIcon       =   "ClassifyLeads.frx":1361
         MousePointer    =   99  'Custom
         Picture         =   "ClassifyLeads.frx":14B3
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   2940
         MouseIcon       =   "ClassifyLeads.frx":180F
         MousePointer    =   99  'Custom
         Picture         =   "ClassifyLeads.frx":1961
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Left            =   2250
         MouseIcon       =   "ClassifyLeads.frx":1C74
         MousePointer    =   99  'Custom
         Picture         =   "ClassifyLeads.frx":1DC6
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   1560
         MouseIcon       =   "ClassifyLeads.frx":20C0
         MousePointer    =   99  'Custom
         Picture         =   "ClassifyLeads.frx":2212
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
         Left            =   870
         MouseIcon       =   "ClassifyLeads.frx":256A
         MousePointer    =   99  'Custom
         Picture         =   "ClassifyLeads.frx":26BC
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4245
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   25
      Top             =   5670
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
         MouseIcon       =   "ClassifyLeads.frx":2A1B
         MousePointer    =   99  'Custom
         Picture         =   "ClassifyLeads.frx":2B6D
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Cancel"
         Top             =   60
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
         MouseIcon       =   "ClassifyLeads.frx":2EAB
         MousePointer    =   99  'Custom
         Picture         =   "ClassifyLeads.frx":2FFD
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmCRIS_ClassifyLeads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsLead                                                            As ADODB.Recordset
Dim AddorEdit                                                         As String

Sub FillSearchGrid(XXX As String)
    Dim TEMPRS                                                        As ADODB.Recordset
    Dim SQL                                                           As String

    lstColor.Enabled = False

    If optCode = True Then
        SQL = "SELECT CODE , LClass, ID  FROM CRIS_LeadClass WHERE CODE LIKE '" & ReplaceQuote(XXX) & "%' ORDER BY ID DESC"
    Else
        SQL = "SELECT CODE , LClass, ID  FROM CRIS_LeadClass WHERE LCLASS LIKE '" & ReplaceQuote(XXX) & "%' ORDER BY ID DESC"
    End If


    Set TEMPRS = gconDMIS.Execute(SQL)

    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        flex_FillListView TEMPRS, lstColor
        'Listview_Loadval Me.lstColor.ListItems, TEMPRS
        lstColor.Enabled = True
    End If



End Sub

Sub rsRefresh()
    Set RsLead = New ADODB.Recordset
    RsLead.Open "SELECT * FROM CRIS_LeadClass ORDER BY ID DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemVars()

    If Not RsLead.EOF And Not RsLead.BOF Then
        TXTCODE = Null2String(RsLead!CODE)
        txtDescription = Null2String(RsLead!LCLASS)
        chkSO.Value = IIf(Null2Bool(RsLead!SO) = True, 1, 0)
        ChkApp = IIf(Null2Bool(RsLead!APPOINTMENT) = True, 1, 0)
        chkEmail = IIf(Null2Bool(RsLead!EMAIL) = True, 1, 0)
        chkLetter = IIf(Null2Bool(RsLead!LETTER) = True, 1, 0)
        chkLO = IIf(Null2Bool(RsLead!LO) = True, 1, 0)
        chkQU = IIf(Null2Bool(RsLead!QU) = True, 1, 0)
        chkSO = IIf(Null2Bool(RsLead!SO) = True, 1, 0)
        chkTD = IIf(Null2Bool(RsLead!TD) = True, 1, 0)
        labid = RsLead!ID
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If

End Sub

Sub InitMemVars()
    TXTCODE = ""
    txtDescription = ""
    chkSO.Value = 0
    chkLO.Value = 0
    chkTD.Value = 0
    chkQU.Value = 0
    chkLetter.Value = 0
    ChkApp.Value = 0
    chkEmail.Value = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (CLASSIFY LEADS)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(labid), "CLASSIFY LEADS")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    AddColumnHeader "Code,Description", lstColor
    ResizeColumnHeader lstColor, "20,78"
    txtSEARCH.Text = vbNullString
    picTop.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    InitMemVars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub lstColor_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstColor
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

Private Sub lstColor_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstColor_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RsLead.MoveFirst
    RsLead.Find ("ID=" & Item.ListSubItems(2).Text)
    StoreMemVars
    On Error Resume Next
    TXTCODE.SetFocus
End Sub

Private Sub optCode_Click()
    If txtSEARCH = "" Then FillSearchGrid (txtSEARCH.Text)
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub optDesc_Click()
    If txtSEARCH = "" Then FillSearchGrid (txtSEARCH.Text)
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)

End Sub

Private Sub txtSEARCH_Change()
    FillSearchGrid (txtSEARCH.Text)
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "CLASSIFY LEADS") = False Then Exit Sub
    On Error GoTo ErrorCode:

    AddorEdit = "ADD"
    picTop.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    optDesc.Enabled = False
    InitMemVars
    lstColor.Enabled = False
    txtSEARCH.Enabled = False
    optCode.Enabled = False
    optDesc.Enabled = False
    On Error Resume Next
    TXTCODE.SetFocus





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()

    picAdds.Visible = True
    picSaves.Visible = False
    picTop.Enabled = False
    lstColor.Enabled = True
    txtSEARCH.Enabled = True
    fraDetails.Enabled = True
    txtSEARCH.Enabled = False
    optCode.Enabled = True
    optDesc.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "CLASSIFY LEADS") = False Then Exit Sub
    On Error GoTo ErrorCode
    If Not RsLead.BOF Or Not RsLead.EOF Then
        If ShowConfirmDelete = True Then
            SQL_STATEMENT = "delete from CRIS_LeadClass where id = " & labid.Caption

            gconDMIS.Execute (SQL_STATEMENT)
            NEW_LogAudit "X", "CLASSIFY LEADS", SQL_STATEMENT, N2Str2Null(labid), "", "Code :" & TXTCODE, "", ""
            ShowDeletedMsg
            LogAudit "X", "LEAD CLASSIFICATIONS", txtDescription
            FillSearchGrid ""
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    rsRefresh
    InitMemVars
    StoreMemVars
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "CLASSIFY LEADS") = False Then Exit Sub
    On Error GoTo ErrorCode:

    AddorEdit = "EDIT"
    picTop.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    fraDetails.Enabled = False




    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next

    txtSEARCH.SetFocus
End Sub

Private Sub cmdNext_Click()
    RsLead.MoveNext
    If RsLead.EOF Then
        RsLead.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    RsLead.MovePrevious
    If RsLead.BOF Then
        RsLead.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    Dim vtxtCode                                                      As String
    Dim vtxtDescription                                               As String
    Dim vchkSO                                                        As Integer
    Dim vchkLO                                                        As Integer
    Dim vchkTD                                                        As Integer
    Dim vchkQU                                                        As Integer
    Dim vchkLetter                                                    As Integer
    Dim vchkApp                                                       As Integer
    Dim vchkEmail                                                     As Integer

    On Error GoTo ErrorCode:

    vchkSO = Null2Bool(chkSO.Value)
    vchkLO = Null2Bool(chkLO.Value)
    vchkTD = Null2Bool(chkTD.Value)
    vchkQU = Null2Bool(chkQU.Value)
    vchkLetter = Null2Bool(chkLetter.Value)
    vchkApp = Null2Bool(ChkApp.Value)
    vchkEmail = Null2Bool(chkEmail.Value)

    Dim lng                                                           As Integer
    If TXTCODE.Text = "" Or txtDescription.Text = "" Then
        ShowIsRequiredMsg "Code and Description"
        On Error Resume Next
        TXTCODE.SetFocus
        Exit Sub
    End If
    '    If (vchkSO = 0 And vchkLO = 0 And vchkTD = 0 And vchkQU = 0 And vchkLetter = 0 And vchkApp = 0 And vchkEmail = 0) Then
    '       chkSO.SetFocus
    '      Exit Sub
    ' End If
    ''''''
    lng = gconDMIS.Execute("select Count(*) from CRIS_LeadClass WHERE CODE=" & N2Str2Null(TXTCODE)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exist"
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(Null2String(RsLead!CODE)) <> UCase(TXTCODE) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exist"
            Exit Sub
        End If
    End If
    If AddorEdit = "ADD" Then
        SQL = "INSERT INTO CRIS_LeadClass(SO, LO, TD, QU, LETTER, EMAIL, APPOINTMENT,  LCLASS,  CODE)  values (" & vbCrLf
        SQL = SQL & vchkSO & ", " & vbCrLf
        SQL = SQL & vchkLO & ", " & vbCrLf
        SQL = SQL & vchkTD & ", " & vbCrLf
        SQL = SQL & vchkQU & ", " & vbCrLf
        SQL = SQL & vchkLetter & ", " & vbCrLf
        SQL = SQL & vchkEmail & ", " & vbCrLf
        SQL = SQL & vchkApp & ", " & vbCrLf
        SQL = SQL & N2Str2Null(txtDescription) & ", " & vbCrLf
        SQL = SQL & N2Str2Null(TXTCODE) & " ) "
        gconDMIS.Execute SQL

        '**********NEW LOG AUDIT****************
        SQL_STATEMENT = SQL
        NEW_LogAudit "A", "CLASSIFY LEADS", SQL_STATEMENT, FindTransactionID(N2Str2Null(TXTCODE), "CODE", "CRIS_LeadClass"), "", "CODE: " & TXTCODE, "", ""
        '**********NEW LOG AUDIT****************
        LogAudit "A", "LEAD CLASSIFICATIONS", txtDescription
    Else
        SQL = "UPDATE CRIS_LeadClass SET " & vbCrLf
        SQL = SQL & "SO=" & vchkSO & ", " & vbCrLf
        SQL = SQL & "LO=" & vchkLO & ", " & vbCrLf
        SQL = SQL & "QU=" & vchkQU & ", " & vbCrLf
        SQL = SQL & "LETTER=" & vchkLetter & ", " & vbCrLf
        SQL = SQL & "EMAIL=" & vchkEmail & ", " & vbCrLf
        SQL = SQL & "APPOINTMENT=" & vchkApp & ", " & vbCrLf
        SQL = SQL & "TD=" & vchkTD & ", " & vbCrLf
        SQL = SQL & "LCLASS=" & N2Str2Null(txtDescription) & ", " & vbCrLf
        SQL = SQL & "CODE=" & N2Str2Null(TXTCODE) & " WHERE ID= " & labid
        gconDMIS.Execute SQL
        '**********NEW LOG AUDIT****************
        SQL_STATEMENT = SQL
        NEW_LogAudit "E", "CLASSIFY LEADS", SQL_STATEMENT, N2Str2Null(labid), "", "CODE: " & TXTCODE, "", ""
        '**********NEW LOG AUDIT****************
        LogAudit "E", "LEAD CLASSIFICATIONS", txtDescription
    End If

    rsRefresh
    If AddorEdit = "EDIT" Then
        RsLead.Find ("ID=" & labid)
    End If
    cmdCancel.Value = True
    FillSearchGrid ""
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If picAdds.Visible = True And KeyCode = vbKeyEscape Then
        Unload Me
    Else
        MoveKeyPress KeyCode
    End If

End Sub


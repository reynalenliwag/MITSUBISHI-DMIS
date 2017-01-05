VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSMIS_Files_PDICheckList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PDI INSPECTION CHECK LIST"
   ClientHeight    =   5355
   ClientLeft      =   75
   ClientTop       =   495
   ClientWidth     =   5835
   ForeColor       =   &H00FFFFFF&
   Icon            =   "PDICheckList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5355
   ScaleWidth      =   5835
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
      Height          =   1095
      Left            =   60
      TabIndex        =   0
      Top             =   -60
      Width           =   5715
      Begin VB.ComboBox cboPDI_Category 
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
         TabIndex        =   4
         Top             =   570
         Width           =   2865
      End
      Begin VB.CheckBox chkPDI_Measurable 
         Caption         =   "Measurable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4110
         TabIndex        =   5
         Top             =   600
         Width           =   1305
      End
      Begin VB.TextBox txtPDI_InspectionName 
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
         MaxLength       =   70
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   180
         Width           =   4440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   450
         TabIndex        =   3
         Top             =   570
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   225
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3375
      Left            =   60
      TabIndex        =   8
      Top             =   990
      Width           =   5715
      Begin VB.OptionButton optCode 
         Caption         =   "&Category"
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
         Left            =   2430
         TabIndex        =   10
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
         Left            =   1080
         TabIndex        =   9
         Top             =   180
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
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   570
         Width           =   5535
      End
      Begin MSComctlLib.ListView lvPDI 
         Height          =   2325
         Left            =   60
         TabIndex        =   14
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
         MouseIcon       =   "PDICheckList.frx":08CA
         NumItems        =   0
      End
      Begin VB.Label labCategory 
         Caption         =   "Label5"
         Height          =   285
         Left            =   3690
         TabIndex        =   12
         Top             =   240
         Width           =   735
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
         TabIndex        =   11
         Top             =   210
         Width           =   1065
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   4320
      ScaleHeight     =   945
      ScaleWidth      =   1800
      TabIndex        =   23
      Top             =   4410
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
         MouseIcon       =   "PDICheckList.frx":0A2C
         MousePointer    =   99  'Custom
         Picture         =   "PDICheckList.frx":0B7E
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Cancel"
         Top             =   45
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
         MouseIcon       =   "PDICheckList.frx":0EBC
         MousePointer    =   99  'Custom
         Picture         =   "PDICheckList.frx":100E
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Save this Record"
         Top             =   45
         Width           =   705
      End
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   30
      ScaleHeight     =   945
      ScaleWidth      =   6075
      TabIndex        =   15
      Top             =   4395
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
         MouseIcon       =   "PDICheckList.frx":135E
         MousePointer    =   99  'Custom
         Picture         =   "PDICheckList.frx":14B0
         Style           =   1  'Graphical
         TabIndex        =   22
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
         MouseIcon       =   "PDICheckList.frx":1816
         MousePointer    =   99  'Custom
         Picture         =   "PDICheckList.frx":1968
         Style           =   1  'Graphical
         TabIndex        =   21
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
         MouseIcon       =   "PDICheckList.frx":1C93
         MousePointer    =   99  'Custom
         Picture         =   "PDICheckList.frx":1DE5
         Style           =   1  'Graphical
         TabIndex        =   20
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
         MouseIcon       =   "PDICheckList.frx":2141
         MousePointer    =   99  'Custom
         Picture         =   "PDICheckList.frx":2293
         Style           =   1  'Graphical
         TabIndex        =   19
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
         MouseIcon       =   "PDICheckList.frx":25A6
         MousePointer    =   99  'Custom
         Picture         =   "PDICheckList.frx":26F8
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
         Left            =   1560
         MouseIcon       =   "PDICheckList.frx":29F2
         MousePointer    =   99  'Custom
         Picture         =   "PDICheckList.frx":2B44
         Style           =   1  'Graphical
         TabIndex        =   17
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
         MouseIcon       =   "PDICheckList.frx":2E9C
         MousePointer    =   99  'Custom
         Picture         =   "PDICheckList.frx":2FEE
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
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
      TabIndex        =   7
      Top             =   690
      Width           =   225
   End
End
Attribute VB_Name = "frmSMIS_Files_PDICheckList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsPDI                                                             As ADODB.Recordset
Dim AddorEdit                                                         As String

Private Function GetModelCode(XXX As String) As String
    Dim rsModelCode                                                   As ADODB.Recordset
    Set rsModelCode = gconDMIS.Execute("select CODE FROM ALL_ModelCode where description=" & N2Str2Null(XXX))
    If Not rsModelCode.EOF Or Not rsModelCode.BOF Then
        GetModelCode = Null2String(rsModelCode!CODE)
    End If
    Set rsModelCode = Nothing
End Function

Private Function SetModel(ModelCode As String) As String
    Dim rsModelCode                                                   As ADODB.Recordset
    Set rsModelCode = gconDMIS.Execute("select description FROM ALL_ModelCode where CODE=" & N2Str2Null(ModelCode))
    If Not rsModelCode.EOF Or Not rsModelCode.BOF Then
        SetModel = Null2String(rsModelCode!Description)
    End If
    Set rsModelCode = Nothing
End Function

Private Function GetCategory(XXX As String) As String
    XXX = UCase(RTrim(LTrim(XXX)))
    Select Case XXX
        Case "VE"
            GetCategory = "VEHICLE EXTERIOR"
        Case "VI"
            GetCategory = "VEHICLE INTERIOR"
        Case "EC"
            GetCategory = "ENGINE COMPARTMENT"
        Case "EE"
            GetCategory = "ELECTRICAL"
        Case "TO"
            GetCategory = "TOOLS"
    End Select
End Function

Private Function SETCATEGORY(ModelCode As String) As String
    ModelCode = UCase(RTrim(LTrim(ModelCode)))
    Select Case UCase(ModelCode)
        Case "VEHICLE EXTERIOR"
            SETCATEGORY = "VE"
        Case "VEHICLE INTERIOR"
            SETCATEGORY = "VI"
        Case "ENGINE COMPARTMENT"
            SETCATEGORY = "EC"
        Case "ELECTRICAL"
            SETCATEGORY = "EE"
        Case "TOOLS"
            SETCATEGORY = "TO"
    End Select
End Function

Sub FillSearchGrid(XXX As String)
    Dim rsPID                                                         As ADODB.Recordset
    Dim SQL                                                           As String

    lvPDI.Sorted = False
    lvPDI.Enabled = False
    lvPDI.ListItems.Clear

    Set rsPID = New ADODB.Recordset


    If optCode.Value = True Then
        Set rsPID = gconDMIS.Execute(SQL & " SELECT INSPECTIONNAME, CATEGORY, PDI_ID FROM SMIS_PDI_LIST WHERE Category like'" & ReplaceQuote(XXX) & "%' order by Category asc")
    ElseIf optDesc.Value = True Then
        Set rsPID = gconDMIS.Execute(SQL & " SELECT INSPECTIONNAME, CATEGORY, PDI_ID FROM SMIS_PDI_LIST where InspectionName like'" & ReplaceQuote(XXX) & "%' order by InspectionName asc")
    End If

    If Not (rsPID.EOF And rsPID.BOF) Then
        Listview_Loadval lvPDI.ListItems, rsPID
        'flex_FillListView rsPID, lvPDI,  True, False
        lvPDI.Refresh
        lvPDI.Enabled = True
    End If

End Sub

Sub InitData()
    With cboPDI_Category
        .AddItem "Vehicle Exterior"
        .AddItem "Vehicle Interior"
        .AddItem "Engine Compartment"
        .AddItem "Electrical"
        .AddItem "Tools"
    End With
End Sub

Sub InitMemVars()
    txtPDI_InspectionName = ""
    cboPDI_Category = ""
    chkPDI_Measurable = 0
    labCategory = ""
End Sub

Sub rsRefresh()
    Set rsPDI = New ADODB.Recordset
    rsPDI.Open "select * from SMIS_PDI_LIST order by PDI_ID DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not rsPDI.EOF And Not rsPDI.BOF Then
        labid.Caption = rsPDI!PDI_ID
        txtPDI_InspectionName = Null2String(rsPDI!InspectionName)
        labCategory = Null2String(rsPDI!Category)
        cboPDI_Category = GetCategory(labCategory)


        If (rsPDI!isunit) = False Then
            chkPDI_Measurable.Value = 0
        Else
            If rsPDI!isunit = True Then
                chkPDI_Measurable.Value = 1
            Else
                chkPDI_Measurable.Value = 0
            End If

        End If
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Private Sub cboPDI_Category_Change()
    If AddorEdit = "" Then: Exit Sub
    labCategory = SETCATEGORY(cboPDI_Category)
End Sub

Private Sub cboPDI_Category_Click()
    cboPDI_Category_Change
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "PDI CHECKLIST") = False Then Exit Sub
    On Error GoTo ErrorCode:

    AddorEdit = "ADD"
    Frame1.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    InitMemVars
    lvPDI.Enabled = False
    txtSEARCH.Enabled = False
    fraDetails.Enabled = False
    fraDetails.Enabled = False
    On Error Resume Next
    txtPDI_InspectionName.SetFocus

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    lvPDI.Enabled = True
    txtSEARCH.Enabled = True
    fraDetails.Enabled = True
    fraDetails.Enabled = True
    fraDetails.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "PDI CHECKLIST") = False Then Exit Sub
    On Error GoTo ErrorCode
    If Not rsPDI.BOF Or Not rsPDI.EOF Then
        If ShowConfirmDelete = True Then
            gconDMIS.Execute "delete from SMIS_PDI_LIST where PDI_ID = " & labid.Caption
            LogAudit "X", "PDI CHECKLIST MASTER FILE", txtPDI_InspectionName
            ShowDeletedMsg
            FillSearchGrid ""
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    rsRefresh
    StoreMemVars
    If FormExist("frmSMIS_Files_PDISetup") Then
        frmSMIS_Files_PDISetup.cboPDI_Model_Change
    End If


    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "PDI CHECKLIST") = False Then Exit Sub
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    fraDetails.Enabled = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub cmdNext_Click()
    On Error GoTo ErrorCode:

    rsPDI.MoveNext
    If rsPDI.EOF Then
        rsPDI.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode:

    rsPDI.MovePrevious
    If rsPDI.BOF Then
        rsPDI.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub cmdSave_Click()
    Dim lng                                                           As Integer
    On Error GoTo ErrorCode:

    If RTrim(Trim(txtPDI_InspectionName)) = "" Then
        ShowIsRequiredMsg "PDI Inspection Description"
        On Error Resume Next
        txtPDI_InspectionName.SetFocus
        Exit Sub
    End If
    If RTrim(Trim(labCategory)) = "" Then
        ShowIsRequiredMsg "Category"
        On Error Resume Next
        cboPDI_Category.SetFocus
        Exit Sub
    End If

    Dim vtxtInspectionName, vtxtCategory, vtxtModelCode               As String
    Dim vtxtIsunit                                                    As Integer
    vtxtInspectionName = N2Str2Null(txtPDI_InspectionName)
    vtxtCategory = N2Str2Null(labCategory)
    vtxtIsunit = chkPDI_Measurable.Value


    Dim rsHanapID                                                     As ADODB.Recordset
    Dim vID                                                           As String
    Set rsHanapID = New ADODB.Recordset
    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "Insert into SMIS_PDI_LIST" & _
                      " (InspectionName,IsUnit,Category)" & _
                      " values (" & vtxtInspectionName & ", " & vtxtIsunit & ", " & vtxtCategory & ")"

        gconDMIS.Execute (SQL_STATEMENT)


        ' THIS IS METHOD IS DONE BECAUSE IN THE DATABASE THERE IS NO Name Called ID
        Set rsHanapID = gconDMIS.Execute("SELECT PDI_ID FROM SMIS_PDI_LIST WHERE InspectionName='" & Null2String(txtPDI_InspectionName) & "'")
        If Not (rsHanapID.BOF And rsHanapID.EOF) Then
            vID = Null2String(rsHanapID!PDI_ID)
        End If

        '*********NEW LOG AUDIT**********
        NEW_LogAudit "A", "PDI CHECKLIST", SQL_STATEMENT, N2Str2Null(vID), "", "INSPECTION NAME:" & txtPDI_InspectionName, "", ""
        '********************************
        LogAudit "A", "PDI CHECKLIST MASTER FILE", txtPDI_InspectionName
    Else
        SQL_STATEMENT = "update SMIS_PDI_LIST set" & _
                      " InspectionName = " & vtxtInspectionName & "," & _
                      " IsUnit = " & vtxtIsunit & "," & _
                      " Category = " & vtxtCategory & _
                      " where PDI_ID = " & labid
        '**************NEW LOG AUDIT***************
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "E", "PDI CHECKLIST", SQL_STATEMENT, N2Str2Null(labid), "", "INSPECTION NAME:" & txtPDI_InspectionName, "", ""
        '**************NEW LOG AUDIT***************

        LogAudit "E", "PDI CHECKLIST MASTER FILE", cboPDI_Category & " " & txtPDI_InspectionName
    End If

    rsRefresh
    If AddorEdit = "EDIT" Then
        rsPDI.Find ("PDI_ID=" & labid)
    End If
    cmdCancel.Value = True
    FillSearchGrid txtSEARCH
    If FormExist("frmSMIS_Files_PDISetup") Then
        frmSMIS_Files_PDISetup.cboPDI_Model_Change
    End If





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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (PDI CHECKLIST)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(labid), "PDI CHECKLIST")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    AddColumnHeader "Description,Category", lvPDI
    ResizeColumnHeader lvPDI, "75,15"
    txtSEARCH = ""
    Frame1.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    InitData
    InitMemVars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub lvPDI_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvPDI
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

Private Sub lvPDI_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lvPDI_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsPDI.MoveFirst
    rsPDI.Find ("PDI_ID=" & Item.ListSubItems(2).Text)
    StoreMemVars
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

Private Sub txtColor_code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPDI_InspectionName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSEARCH_Change()
    FillSearchGrid (txtSEARCH.Text)
End Sub

Public Sub SearchID(Lid)
    rsPDI.MoveFirst
    rsPDI.Find ("PDI_ID=" & Lid)
    StoreMemVars
    If Not (rsPDI.EOF Or rsPDI.BOF) Then
        cmdEdit.Value = True
    End If
End Sub


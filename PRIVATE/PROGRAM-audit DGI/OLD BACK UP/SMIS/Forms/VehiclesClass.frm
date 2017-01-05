VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSMIS_Files_VehiclesClass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Class"
   ClientHeight    =   5325
   ClientLeft      =   75
   ClientTop       =   435
   ClientWidth     =   5835
   ForeColor       =   &H00FFFFFF&
   Icon            =   "VehiclesClass.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5325
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
      Height          =   1035
      Left            =   30
      TabIndex        =   2
      Top             =   -60
      Width           =   5715
      Begin VB.TextBox txtClassCode 
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
         Left            =   1275
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   180
         Width           =   1200
      End
      Begin VB.TextBox txtVehicleClass 
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
         Left            =   1275
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   4365
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
         Left            =   705
         TabIndex        =   4
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Class"
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
      Begin MSComctlLib.ListView lstClass 
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
         MouseIcon       =   "VehiclesClass.frx":08CA
         NumItems        =   0
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
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   6075
      TabIndex        =   13
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
         MouseIcon       =   "VehiclesClass.frx":0A2C
         MousePointer    =   99  'Custom
         Picture         =   "VehiclesClass.frx":0B7E
         Style           =   1  'Graphical
         TabIndex        =   14
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
         MouseIcon       =   "VehiclesClass.frx":0EE4
         MousePointer    =   99  'Custom
         Picture         =   "VehiclesClass.frx":1036
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
         Left            =   3630
         MouseIcon       =   "VehiclesClass.frx":1361
         MousePointer    =   99  'Custom
         Picture         =   "VehiclesClass.frx":14B3
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
         Left            =   2940
         MouseIcon       =   "VehiclesClass.frx":180F
         MousePointer    =   99  'Custom
         Picture         =   "VehiclesClass.frx":1961
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
         Left            =   2250
         MouseIcon       =   "VehiclesClass.frx":1C74
         MousePointer    =   99  'Custom
         Picture         =   "VehiclesClass.frx":1DC6
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
         MouseIcon       =   "VehiclesClass.frx":20C0
         MousePointer    =   99  'Custom
         Picture         =   "VehiclesClass.frx":2212
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
         MouseIcon       =   "VehiclesClass.frx":256A
         MousePointer    =   99  'Custom
         Picture         =   "VehiclesClass.frx":26BC
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4275
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   21
      Top             =   4425
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
         MouseIcon       =   "VehiclesClass.frx":2A1B
         MousePointer    =   99  'Custom
         Picture         =   "VehiclesClass.frx":2B6D
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   60
         MouseIcon       =   "VehiclesClass.frx":2EAB
         MousePointer    =   99  'Custom
         Picture         =   "VehiclesClass.frx":2FFD
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Save Vehicle Class"
         Top             =   45
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
Attribute VB_Name = "frmSMIS_Files_VehiclesClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsClass                                                           As ADODB.Recordset
Dim AddorEdit                                                         As String

Sub FillSearchGrid(XXX As String)
    Dim rsClass2                                                      As ADODB.Recordset
    lstClass.Sorted = False
    lstClass.Enabled = False
    lstClass.ListItems.Clear
    Set rsClass2 = New ADODB.Recordset

    If optCode.Value = True Then
        Set rsClass2 = gconDMIS.Execute("select  CODE , CLASSNAME, ID from SMIS_VEHICLESCLASS where CODE like'" & ReplaceQuote(XXX) & "%' order by CODE asc")
    Else
        Set rsClass2 = gconDMIS.Execute("select  CODE , CLASSNAME, ID from SMIS_VEHICLESCLASS where CLASSNAME like'" & ReplaceQuote(XXX) & "%' order by CLASSNAME asc")
    End If

    If Not (rsClass2.EOF And rsClass2.BOF) Then
        Listview_Loadval Me.lstClass.ListItems, rsClass2
        lstClass.Refresh
        lstClass.Enabled = True
    End If

End Sub

Sub InitMemVars()
    txtClassCode.Text = vbNullString
    txtVehicleClass.Text = vbNullString
    labid = 0
End Sub

Sub rsRefresh()
    Set rsClass = New ADODB.Recordset
    rsClass.Open "select * from SMIS_VehiclesClass order by id DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not rsClass.EOF And Not rsClass.BOF Then
        labid.Caption = rsClass!ID
        txtClassCode.Text = Null2String(rsClass!CODE)
        txtVehicleClass.Text = Null2String(rsClass!ClassName)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "VEHICLE CLASS") = False Then Exit Sub
    On Error GoTo ErrorCode:

    AddorEdit = "ADD"
    Frame1.Enabled = True
    fraDetails.Enabled = False
    picAdds.Visible = False
    picSaves.Visible = True
    InitMemVars
    lstClass.Enabled = False
    txtSEARCH.Enabled = False
    On Error Resume Next
    txtClassCode.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    fraDetails.Enabled = True
    picAdds.Visible = True
    picSaves.Visible = False
    lstClass.Enabled = True
    txtSEARCH.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "VEHICLE CLASS") = False Then Exit Sub
    On Error GoTo ErrorCode

    If gconDMIS.Execute("SELECT Count(*) FROM SMIS_MrrInv where Class=" & N2Str2Null(rsClass!CODE)).Fields(0).Value > 0 Then
        MessagePop RecLocekd, "Cannot Edit or Delete", "Current Class Name and Code is in use .. Cannot Delete"
        Exit Sub
    End If
    
    If Not rsClass.BOF Or Not rsClass.EOF Then
        If ShowConfirmDelete = True Then
            SQL_STATEMENT = "delete from SMIS_VehiclesClass where id = " & labid.Caption
            gconDMIS.Execute (SQL_STATEMENT)
            
            NEW_LogAudit "X", "VEHICLE CLASS", SQL_STATEMENT, labid, "", "CODE :" & txtClassCode, "", ""
            
            ShowDeletedMsg
            FillSearchGrid ""
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    
    rsRefresh
    StoreMemVars
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "VEHICLE CLASS") = False Then Exit Sub
    On Error GoTo ErrorCode:

    AddorEdit = "EDIT"
    Frame1.Enabled = True
    fraDetails.Enabled = False
    picAdds.Visible = False
    picSaves.Visible = True

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
    rsClass.MoveNext
    If rsClass.EOF Then
        rsClass.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsClass.MovePrevious
    If rsClass.BOF Then
        rsClass.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub cmdSave_Click()
    Dim lng                                                           As Integer
    On Error GoTo ErrorCode:

    If txtClassCode.Text = "" Or txtVehicleClass.Text = "" Then
        ShowIsRequiredMsg "Vehicles Class Code and Class Name"
        On Error Resume Next
        txtClassCode.SetFocus
        Exit Sub
    End If
    ''''''
    lng = gconDMIS.Execute("select Count(*) from SMIS_VEHICLESCLASS WHERE CODE=" & N2Str2Null(txtClassCode)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exist"
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(Null2String(rsClass!CODE)) <> UCase(txtClassCode) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exist"
            Exit Sub
        End If
    End If

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "Insert into SMIS_VEHICLESCLASS" & _
                      " (CODE,ClassName)" & _
                      " values (" & N2Str2Null(txtClassCode.Text) & ", " & N2Str2Null(txtVehicleClass.Text) & ")"

        '************NEW LOG AUDIT******************
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "A", "VEHICLE CLASS", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtClassCode), "CODE", "SMIS_VEHICLESCLASS"), "", "CODE :" & txtClassCode, "", ""
        '************NEW LOG AUDIT******************
        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update SMIS_VEHICLESCLASS set" & _
                      " CODE = " & N2Str2Null(txtClassCode.Text) & "," & _
                      " ClassName = " & N2Str2Null(txtVehicleClass.Text) & _
                      " where id = " & labid.Caption
        '************NEW LOG AUDIT******************
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "E", "VEHICLE CLASS", SQL_STATEMENT, N2Str2Null(labid), "", "CODE :" & txtClassCode, "", ""
        '************NEW LOG AUDIT******************
        ShowSuccessFullyUpdated
    End If
    rsRefresh
    If AddorEdit = "EDIT" Then
        rsClass.Find "id =" & labid.Caption
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (VEHICLE CLASS)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(labid), "VEHICLE CLASS")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    AddColumnHeader "Code,Class", lstClass
    ResizeColumnHeader lstClass, "20,78"

    txtSEARCH.Text = vbNullString
    Frame1.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    InitMemVars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub lstClass_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstClass
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

Private Sub lstClass_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstClass_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsClass.MoveFirst
    rsClass.Find ("ID=" & Item.ListSubItems(2).Text)
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

Private Sub txtClassCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSEARCH_Change()
    FillSearchGrid (txtSEARCH.Text)
End Sub

Private Sub txtVehicleClass_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub


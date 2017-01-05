VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSTechnician 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Technician Master File"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Technician.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8955
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   2745
      ScaleHeight     =   945
      ScaleWidth      =   6435
      TabIndex        =   12
      Top             =   3960
      Width           =   6435
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
         MouseIcon       =   "Technician.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "Technician.frx":0A1C
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
         MouseIcon       =   "Technician.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "Technician.frx":0ED4
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
         MouseIcon       =   "Technician.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "Technician.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Delete Selected Technician"
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
         MouseIcon       =   "Technician.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "Technician.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Edit Selected Technician"
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
         MouseIcon       =   "Technician.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "Technician.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Add Technician"
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
         MouseIcon       =   "Technician.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "Technician.frx":211C
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
         MouseIcon       =   "Technician.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "Technician.frx":2568
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
         MouseIcon       =   "Technician.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "Technician.frx":2A12
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
      Height          =   3915
      Left            =   2730
      TabIndex        =   3
      Top             =   30
      Width           =   6105
      Begin VB.TextBox txtTech_Name 
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
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   1
         Top             =   630
         Width           =   4305
      End
      Begin VB.TextBox txtSkills 
         Appearance      =   0  'Flat
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
         Height          =   2535
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Text            =   "Technician.frx":2D71
         Top             =   1290
         Width           =   5925
      End
      Begin VB.TextBox txtTechnician 
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
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Technician Name"
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
         Left            =   90
         TabIndex        =   11
         Top             =   660
         Width           =   1605
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Technician Code"
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
         Left            =   90
         TabIndex        =   7
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Technican Skills"
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
         Left            =   90
         TabIndex        =   4
         Top             =   1020
         Width           =   2685
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4845
      Left            =   60
      TabIndex        =   8
      Top             =   30
      Width           =   2625
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
         Left            =   60
         MaxLength       =   35
         TabIndex        =   9
         Top             =   150
         Width           =   2505
      End
      Begin MSComctlLib.ListView lstTechnician 
         Height          =   4245
         Left            =   30
         TabIndex        =   10
         Top             =   540
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   7488
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
         MouseIcon       =   "Technician.frx":2D76
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   1411
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
      Left            =   7380
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   21
      Top             =   3990
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
         Left            =   690
         MouseIcon       =   "Technician.frx":2ED8
         MousePointer    =   99  'Custom
         Picture         =   "Technician.frx":302A
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
         Left            =   0
         MouseIcon       =   "Technician.frx":3368
         MousePointer    =   99  'Custom
         Picture         =   "Technician.frx":34BA
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Save Technician"
         Top             =   30
         Width           =   705
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
Attribute VB_Name = "frmCSMSTechnician"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTechnician                                                      As ADODB.Recordset
Dim AddorEdit                                                         As String

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "TECHNICIAN") = False Then Exit Sub
    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False: Picture2.Visible = True
    Picture1.Enabled = False: Picture2.Enabled = True
    InitMemvars
    On Error Resume Next
    txtTechnician.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True: Picture2.Visible = False
    Picture1.Enabled = True: Picture2.Enabled = False
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "TECHNICIAN") = False Then Exit Sub
    On Error GoTo Errorcode
    If Not rsTechnician.BOF Or Not rsTechnician.EOF Then
        If ShowConfirmDelete = True Then
             gconDMIS.Execute "delete from CSMS_Technicians where id = " & labid.Caption
             LogAudit "X", "TECHNICIAN DETAILS", txtTech_Name
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
    If Function_Access(LOGID, "Acess_EDIT", "TECHNICIAN") = False Then Exit Sub
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False: Picture2.Visible = True
    Picture1.Enabled = False: Picture2.Enabled = True
    On Error Resume Next
    txtTechnician.SetFocus
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
    rsTechnician.MoveNext
    If rsTechnician.EOF Then
        rsTechnician.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsTechnician.MovePrevious
    If rsTechnician.BOF Then
        rsTechnician.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "TECHNICIAN") = False Then Exit Sub
LogAudit "V", "TECHNICIAN REPORT"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim VTXTTechnician, VTXTTech_Name, VTXTSkills                     As String

    If txtTech_Name.Text = "" Then
        ShowIsRequiredMsg "Technician Name"
        On Error Resume Next
        txtTech_Name.SetFocus
        Exit Sub
    End If
    If IsNull(txtTechnician.Text) = True Then
        ShowIsRequiredMsg "Technician Code"
        On Error Resume Next
        txtTechnician.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Dim rsfindDup                                             As ADODB.Recordset
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select Technician from CSMS_vw_Technician where Technician = '" & txtTechnician.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "Technician Code already exist!"
                On Error Resume Next
                txtTechnician.SetFocus
                Exit Sub
            End If
        End If
    End If

    VTXTTechnician = N2Str2Null(txtTechnician.Text)
    VTXTTech_Name = N2Str2Null(txtTech_Name.Text)
    VTXTSkills = N2Str2Null(txtSkills.Text)

    If AddorEdit = "ADD" Then
        Dim rsGetTechEmpNo                                            As ADODB.Recordset
        Dim NewEmpNo                                                  As String
        Set rsGetTechEmpNo = New ADODB.Recordset
        Set rsGetTechEmpNo = gconDMIS.Execute("select * from CSMS_Technicians Order by Technician asc")
        If Not rsGetTechEmpNo.EOF And Not rsGetTechEmpNo.BOF Then
            NewEmpNo = "'T" & Format(NumericVal(Right(rsGetTechEmpNo!empno, 3)) + 1, "000") & "'"
        Else
            NewEmpNo = "'T001'"
        End If
        gconDMIS.Execute "Insert into CSMS_Technicians " & _
                         "(EmpNo,Technician,Tech_Name,FirstName,LastUpdate,UserCode)" & _
                       " values (" & NewEmpNo & "," & _
                         VTXTTechnician & ", " & VTXTTech_Name & ", " & VTXTTech_Name & ", " & N2Str2Null(LOGDATE) & _
                         ", " & N2Str2Null(LOGCODE) & ")"
        LogAudit "A", "TECHNICIAN DETAILS", txtTech_Name
    
    Else
        gconDMIS.Execute "update CSMS_Technicians set" & _
                       " Technician = " & VTXTTechnician & "," & _
                       " Tech_Name = " & VTXTTech_Name & "," & _
                       " FirstName = " & VTXTTech_Name & "," & _
                       " LastUpdate = " & N2Str2Null(LOGDATE) & "," & _
                       " UserCode = " & N2Str2Null(LOGCODE) & _
                       " where id = " & labid.Caption
        LogAudit "E", "TECHNICIAN DETAILS", txtTech_Name
    End If
    rsRefresh
    On Error Resume Next
    rsTechnician.Find "Technician = " & VTXTTechnician
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
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    Frame1.Enabled = False: textSearch.Text = ""
    InitMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub InitMemvars()
    txtTechnician.Text = ""
    txtSkills.Text = ""
    txtTech_Name.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsTechnician.EOF And Not rsTechnician.BOF Then
        labid.Caption = rsTechnician!ID
        txtTechnician.Text = Null2String(rsTechnician!TECHNICIAN)
        txtTech_Name.Text = Null2String(rsTechnician!Tech_Name)
        txtSkills.Text = ShowTechSkills(Null2String(rsTechnician!empno))
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set rsTechnician = New ADODB.Recordset
    rsTechnician.Open "select * from CSMS_vw_Technician order by Technician asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Function ShowTechSkills(XXX As String) As String
    Dim rsTechSkills                                                  As ADODB.Recordset
    Set rsTechSkills = New ADODB.Recordset
    Set rsTechSkills = gconDMIS.Execute("Select * from CSMS_vw_TECHSKILLS where EmpNo = '" & XXX & "'")
    ShowTechSkills = ""
    If Not rsTechSkills.EOF And Not rsTechSkills.BOF Then
        rsTechSkills.MoveFirst
        Do While Not rsTechSkills.EOF
            ShowTechSkills = ShowTechSkills + Null2String(rsTechSkills!SKILLS) & vbCrLf
            rsTechSkills.MoveNext
        Loop
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub lstTechnician_GotFocus()
    rsTechnician.Bookmark = rsFind(rsTechnician.Clone, "Technician", lstTechnician.SelectedItem).Bookmark
    StoreMemVars
End Sub

Private Sub lstTechnician_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsTechnician.Bookmark = rsFind(rsTechnician.Clone, "Technician", lstTechnician.SelectedItem).Bookmark
    StoreMemVars
End Sub

Private Sub lstTechnician_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstTechnician
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

Private Sub lstTechnician_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstTechnician_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then On Error Resume Next: textSearch.SetFocus
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (textSearch.Text)
    End If
End Sub

Sub FillGrid()
    Dim rsTechnician2                                                 As ADODB.Recordset
    lstTechnician.Enabled = False
    lstTechnician.Sorted = False: lstTechnician.ListItems.Clear: lstTechnician.Enabled = True
    Set rsTechnician2 = New ADODB.Recordset
    Set rsTechnician2 = gconDMIS.Execute("select Technician, Tech_Name, ID from CSMS_vw_Technician order by Technician asc")
    If Not (rsTechnician2.EOF And rsTechnician2.BOF) Then
        Listview_Loadval Me.lstTechnician.ListItems, rsTechnician2
        lstTechnician.Refresh
        lstTechnician.Enabled = True
    Else
        lstTechnician.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsTechnician2                                                 As ADODB.Recordset
    lstTechnician.Enabled = False
    lstTechnician.Sorted = False: lstTechnician.ListItems.Clear
    Set rsTechnician2 = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsTechnician2 = gconDMIS.Execute("select Technician, Tech_Name, ID from CSMS_vw_Technician where Technician like '" & XXX & "%'")
    If Not (rsTechnician2.EOF And rsTechnician2.BOF) Then
        Listview_Loadval Me.lstTechnician.ListItems, rsTechnician2
        lstTechnician.Refresh
        lstTechnician.Enabled = True
    End If

End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstTechnician.Enabled = True Then
            lstTechnician.SetFocus
        End If
    End If
End Sub

Private Sub txtTech_Name_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

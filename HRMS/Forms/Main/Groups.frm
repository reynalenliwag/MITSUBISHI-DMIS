VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRMSGroups 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Group Codes"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7950
   ForeColor       =   &H00D8E9EC&
   Icon            =   "Groups.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7950
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2295
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   10
      Top             =   3870
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
         MouseIcon       =   "Groups.frx":0442
         MousePointer    =   99  'Custom
         Picture         =   "Groups.frx":0594
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
         MouseIcon       =   "Groups.frx":08FA
         MousePointer    =   99  'Custom
         Picture         =   "Groups.frx":0A4C
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
         MouseIcon       =   "Groups.frx":0DB2
         MousePointer    =   99  'Custom
         Picture         =   "Groups.frx":0F04
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
         MouseIcon       =   "Groups.frx":122F
         MousePointer    =   99  'Custom
         Picture         =   "Groups.frx":1381
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
         MouseIcon       =   "Groups.frx":16DD
         MousePointer    =   99  'Custom
         Picture         =   "Groups.frx":182F
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
         MouseIcon       =   "Groups.frx":1B42
         MousePointer    =   99  'Custom
         Picture         =   "Groups.frx":1C94
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
         MouseIcon       =   "Groups.frx":1F8E
         MousePointer    =   99  'Custom
         Picture         =   "Groups.frx":20E0
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
         MouseIcon       =   "Groups.frx":2438
         MousePointer    =   99  'Custom
         Picture         =   "Groups.frx":258A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   6435
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   19
      Top             =   3870
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
         MouseIcon       =   "Groups.frx":28E9
         MousePointer    =   99  'Custom
         Picture         =   "Groups.frx":2A3B
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
         MouseIcon       =   "Groups.frx":2D79
         MousePointer    =   99  'Custom
         Picture         =   "Groups.frx":2ECB
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   4560
      Left            =   30
      ScaleHeight     =   4500
      ScaleWidth      =   1845
      TabIndex        =   9
      Top             =   180
      Width           =   1905
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   6960
         Left            =   0
         Picture         =   "Groups.frx":321B
         Top             =   0
         Width           =   9915
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   2685
      Left            =   2010
      ScaleHeight     =   2685
      ScaleWidth      =   5865
      TabIndex        =   7
      Top             =   1170
      Width           =   5865
      Begin MSComctlLib.ListView lstCodes_Group 
         Height          =   2565
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   5775
         _ExtentX        =   10186
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
         MouseIcon       =   "Groups.frx":16F78
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
   Begin VB.PictureBox picCodes_Group 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   2010
      ScaleHeight     =   975
      ScaleWidth      =   5865
      TabIndex        =   0
      Top             =   180
      Width           =   5865
      Begin VB.TextBox txtCodes 
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
         TabIndex        =   4
         Top             =   90
         Width           =   765
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
         TabIndex        =   3
         Top             =   510
         Width           =   4425
      End
      Begin Crystal.CrystalReport rptDepartment 
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
         Left            =   -60
         TabIndex        =   6
         Top             =   120
         Width           =   675
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
         Top             =   600
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   570
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmHRMSGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCodes_Group                            As ADODB.Recordset
Dim AddorEdit                                As String
'LAST UPDATE : MARCH 17,2007 : INVALID OBJECT NAME "HMRS_CODES_GROUP" : 20070317 ( MAKOY )
'LAST UPDATE : MARCH 17,2007 : INVALID OBJECT NAME "HRMS_CODES_GROUP" : 20071703 ( MAKOY )
'LAST UPDATE : MARCH 17,2007 : INFINITE LOOP : 20071703 ( MAKOY )

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "FILES GROUPS") = False Then Exit Sub
    AddorEdit = "ADD"
    InitMemVars
    Picture1.Visible = False
    Picture2.Visible = True
End Sub

Private Sub cmdCancel_Click()
    picCodes_Group.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    STOREMEMVARS
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "FILES GROUPS") = False Then Exit Sub
    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from HRMS_Codes_Group where id = " & labID.Caption
        ShowDeletedMsg
    End If
    RsRefresh
    STOREMEMVARS
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "FILES GROUPS") = False Then Exit Sub
    AddorEdit = "EDIT"
    picCodes_Group.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
End Sub

Private Sub cmdExit_Click()
    UnloadForm Me
End Sub

Private Sub cmdFind_Click()
    MsgBox "Pls use the List view to find...", vbInformation, "Find"
End Sub

Private Sub cmdNext_Click()
    rsCodes_Group.MoveNext
    If rsCodes_Group.EOF Then
        rsCodes_Group.MoveLast
        ShowLastRecordMsg
    End If
    STOREMEMVARS
End Sub

Private Sub cmdPrevious_Click()
    rsCodes_Group.MovePrevious
    If rsCodes_Group.BOF Then
        rsCodes_Group.MoveFirst
        ShowFirstRecordMsg
    End If
    STOREMEMVARS
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "FILES GROUPS") = False Then Exit Sub
    Screen.MousePointer = 11
    
    'rptCodes_Group.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    'rptCodes_Group.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    'rptCodes_Group.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    'PrintSQLReport rptCodes_Group, HRMS_REPORT_PATH & "Codes_Group.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    txtCodes.Text = N2Str2Null(txtCodes.Text)
    txtDescription.Text = N2Str2Null(txtDescription.Text)
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_Codes_Group " & _
                         "(Codes,Description) " & _
                       " values (" & txtCodes.Text & ", " & _
                         "" & txtDescription.Text & ")"
        ShowSuccessFullyAdded
    Else
        gconDMIS.Execute "update HRMS_Codes_Group set" & _
                       " Codes = " & txtCodes.Text & "," & _
                       " Description = " & txtDescription.Text & _
                       " where id = " & labID.Caption
        ShowSuccessFullyUpdated
    End If
    RsRefresh
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

    RsRefresh
    STOREMEMVARS
    FillGrid
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Sub RsRefresh()
    Set rsCodes_Group = New ADODB.Recordset
    'LAST UPDATE MARCH 17,2007 : INVALID OBJECT NAME "HRMS_CODES_GROUP" : 20071703 ( MAKOY )
    'rsCodes_Group.Open "select * from HRMS_Codes_Group order by Codes", gconDMIS, adOpenForwardOnly, adLockReadOnly
    rsCodes_Group.Open "select * from HRMS_Codes_Adjustment order by Codes", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub InitMemVars()
    picCodes_Group.Enabled = True
    txtCodes.Text = ""
    txtDescription.Text = ""
End Sub

Sub STOREMEMVARS()
    If Not rsCodes_Group.EOF And Not rsCodes_Group.BOF Then
        picCodes_Group.Enabled = False
        labID.Caption = rsCodes_Group!ID
        txtCodes.Text = Null2String(rsCodes_Group!Codes)
        txtDescription.Text = Null2String(rsCodes_Group!Description)
    Else
        ShowNoRecord
        picCodes_Group.Enabled = False
        'LAST UPDATE : MARCH 17,2007 : INFINITE LOOP : 20071703 ( MAKOY )
        'If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdADD.Value = True Else Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Sub FillGrid()
    Dim rsCodes_Group2                       As ADODB.Recordset
    lstCodes_Group.Sorted = False: lstCodes_Group.ListItems.Clear
    lstCodes_Group.Enabled = False
    Set rsCodes_Group2 = New ADODB.Recordset
    
    'Set rsCodes_Group2 = gconDMIS.Execute("select Codes,Description,ID from HRMS_Codes_Group")
    'LAST UPDATE : MARCH 17,2007 : INVALID OBJECT NAME "HMRS_CODES_GROUP" : 20070317 ( MAKOY )
    
    Set rsCodes_Group2 = gconDMIS.Execute("select Codes,Description,ID from HRMS_Codes_Adjustment")
    If Not (rsCodes_Group2.EOF And rsCodes_Group2.BOF) Then
        Listview_Loadval Me.lstCodes_Group.ListItems, rsCodes_Group2
        lstCodes_Group.Refresh
        lstCodes_Group.Enabled = True
    End If
    
End Sub

Private Sub lstCodes_Group_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstCodes_Group
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

Private Sub lstCodes_Group_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstCodes_Group_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsCodes_Group.Bookmark = rsFind(rsCodes_Group.Clone, "Codes", Me.lstCodes_Group.SelectedItem).Bookmark
    STOREMEMVARS
End Sub

Private Sub txtCodes_LostFocus()
    txtCodes.Text = UCase(txtCodes.Text)
End Sub

VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSMIS_Files_SalesAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Account Executives"
   ClientHeight    =   6390
   ClientLeft      =   210
   ClientTop       =   540
   ClientWidth     =   5790
   ForeColor       =   &H00FCFCFC&
   Icon            =   "SalesAgent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   5790
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   -1440
      ScaleHeight     =   900
      ScaleWidth      =   9225
      TabIndex        =   16
      Top             =   5490
      Width           =   9225
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
         Left            =   6480
         MouseIcon       =   "SalesAgent.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "SalesAgent.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   5790
         MouseIcon       =   "SalesAgent.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "SalesAgent.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Left            =   5100
         MouseIcon       =   "SalesAgent.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "SalesAgent.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   4410
         MouseIcon       =   "SalesAgent.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "SalesAgent.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Left            =   3720
         MouseIcon       =   "SalesAgent.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "SalesAgent.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   3030
         MouseIcon       =   "SalesAgent.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "SalesAgent.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Left            =   2340
         MouseIcon       =   "SalesAgent.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "SalesAgent.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Left            =   1650
         MouseIcon       =   "SalesAgent.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "SalesAgent.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   0
      TabIndex        =   4
      Top             =   30
      Width           =   5730
      Begin Crystal.CrystalReport rptSAE 
         Left            =   4365
         Top             =   1215
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1350
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   1800
         Width           =   2445
      End
      Begin VB.TextBox txtName 
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1350
         TabIndex        =   0
         Top             =   225
         Width           =   4245
      End
      Begin VB.TextBox txtLastName 
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1350
         TabIndex        =   1
         Top             =   630
         Width           =   2475
      End
      Begin VB.TextBox txtFirstName 
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1350
         TabIndex        =   2
         Top             =   1020
         Width           =   2475
      End
      Begin VB.TextBox txtMiddleName 
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1350
         TabIndex        =   3
         Top             =   1410
         Width           =   2475
      End
      Begin VB.Label labSAECODE 
         Height          =   375
         Left            =   4140
         TabIndex        =   28
         Top             =   750
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Team"
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
         Height          =   255
         Left            =   90
         TabIndex        =   14
         Top             =   1845
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Height          =   255
         Left            =   750
         TabIndex        =   10
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
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
         Height          =   255
         Left            =   270
         TabIndex        =   9
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
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
         Height          =   255
         Left            =   270
         TabIndex        =   8
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
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
         Height          =   255
         Left            =   90
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3105
      Left            =   0
      TabIndex        =   11
      Top             =   2340
      Width           =   5775
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
         Height          =   360
         Left            =   90
         MaxLength       =   35
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   150
         Width           =   5595
      End
      Begin MSComctlLib.ListView lstExecutive 
         Height          =   2505
         Left            =   60
         TabIndex        =   13
         Top             =   540
         Width           =   5670
         _ExtentX        =   10001
         _ExtentY        =   4419
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
         MouseIcon       =   "SalesAgent.frx":2D71
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FULL NAME"
            Object.Width           =   176
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Team"
            Object.Width           =   176
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
      Left            =   4275
      ScaleHeight     =   885
      ScaleWidth      =   2580
      TabIndex        =   25
      Top             =   5490
      Width           =   2580
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
         MouseIcon       =   "SalesAgent.frx":2ED3
         MousePointer    =   99  'Custom
         Picture         =   "SalesAgent.frx":3025
         Style           =   1  'Graphical
         TabIndex        =   27
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
         MouseIcon       =   "SalesAgent.frx":3363
         MousePointer    =   99  'Custom
         Picture         =   "SalesAgent.frx":34B5
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Save this Record"
         Top             =   45
         Width           =   705
      End
   End
   Begin VB.Label labPrev 
      Caption         =   "Label4"
      Height          =   315
      Left            =   5160
      TabIndex        =   6
      Top             =   1470
      Width           =   195
   End
   Begin VB.Label labid 
      Caption         =   "Label4"
      Height          =   255
      Left            =   5190
      TabIndex        =   5
      Top             =   1530
      Width           =   225
   End
End
Attribute VB_Name = "frmSMIS_Files_SalesAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSREP                                                            As ADODB.Recordset
Dim AddorEdit                                                         As String

'Upating Code       : AXP-0707200712:28
Private Sub cmdADD_Click()
    If Function_Access(LOGID, "Acess_ADD", "SALES ACCOUNT EXECUTIVE") = False Then Exit Sub
    On Error GoTo ErrorCode:

    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    lstExecutive.Enabled = False
    txtSearch.Enabled = False
    On Error Resume Next
    txtname.SetFocus

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstExecutive.Enabled = True
    txtSearch.Enabled = True
    fraDetails.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "SALES ACCOUNT EXECUTIVE") = False Then Exit Sub
    On Error GoTo ErrorCode
    If Not rsSREP.BOF Or Not rsSREP.EOF Then
        If ShowConfirmDelete = True Then
            gconDMIS.Execute "delete from SMIS_vw_SRep where id = " & LabID.Caption
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

'Upating Code       : AXP-0707200712:28
Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "SALES ACCOUNT EXECUTIVE") = False Then Exit Sub
    On Error GoTo ErrorCode:

    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
    On Error Resume Next
    txtname.SetFocus





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200712:28
Private Sub cmdFind_Click()
    On Error Resume Next

    txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsSREP.MoveNext
    If rsSREP.EOF Then
        rsSREP.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsSREP.MovePrevious
    If rsSREP.BOF Then
        rsSREP.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

'Upating Code       : AXP-0707200712:27
Private Sub cmdPrint_Click()

    If Function_Access(LOGID, "Acess_PRINT", "SALES ACCOUNT EXECUTIVE") = False Then Exit Sub

    On Error GoTo ErrorCode:
    Screen.MousePointer = 11
    With frmMain.rptMain
        .ReportTitle = "SALES EXECITIVE LISTING"
        .Formulas(0) = "CompanyName = '" & Company_name & "'"
        .Formulas(1) = "CompanyAddress = '" & Company_Address & "'"
        .Connect = DMIS_REPORT_Connection
        .WindowTitle = "SALES PERSONNEL LIST"
        .ReportFileName = SMIS_REPORT_PATH & "allsaelist.rpt"
        .Action = 1
        Screen.MousePointer = 0
    End With





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200712:28
Private Sub cmdSave_Click()

    On Error GoTo ErrorCode:

    If txtLastName.Text = "" Or txtFirstName.Text = "" Then
        ShowIsRequiredMsg "Code and Description"
        Exit Sub
    End If
    Dim VTXTName                                                      As String
    Dim VTXTLASTNAME                                                  As String
    Dim VTXTFIRSTNAME                                                 As String
    Dim VTXTMIDDLENAME                                                As String
    Dim VTXTTEAMNAME                                                  As String

    VTXTName = N2Str2Null(txtname.Text)
    VTXTLASTNAME = N2Str2Null(txtLastName.Text)
    VTXTFIRSTNAME = N2Str2Null(txtFirstName.Text)
    VTXTMIDDLENAME = N2Str2Null(txtMiddleName.Text)
    VTXTTEAMNAME = N2Str2Null(Combo1)

    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into SMIS_vw_SRep" & _
                       " (name,fname,lname,middle,teamname)" & _
                       " values (" & VTXTName & ", " & VTXTFIRSTNAME & ", " & VTXTLASTNAME & ", " & VTXTMIDDLENAME & ", " & VTXTTEAMNAME & ")"


    Else



        gconDMIS.Execute " update HRMS_EmpInfo set" & _
                       " Firstname = " & VTXTFIRSTNAME & "," & _
                       " Lastname = " & VTXTLASTNAME & "," & _
                       " Middlename = " & VTXTMIDDLENAME & _
                       " where id = " & LabID.Caption

        gconDMIS.Execute " DELETE FROM SMIS_SALESTEAM WHERE SAEID=" & LabID

        gconDMIS.Execute " INSERT INTO SMIS_SALESTEAM (TEAMNAME, SAEID, SAECODE ) VALUES " & _
                       " ( " & VTXTTEAMNAME & "," & LabID & ", " & N2Str2Null(labSAECODE) & " )"



    End If

    rsRefresh

    Call FillCombo("SELECT DISTINCT TEAMNAME  from SMIS_vw_Srep WHERE LEN(TEAMNAME)>0", -1, 0, Combo1)

    If AddorEdit = "EDIT" Then
        rsSREP.Find "id =" & LabID.Caption
    End If



    cmdCancel.Value = True
    FillGrid





    Exit Sub
ErrorCode:
    ShowVBError

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()

    CenterMe frmMain, Me, 1
    rsRefresh
    txtSearch.Text = ""
    Frame1.Enabled = False
    initMemvars
    Picture1.Visible = True
    Picture2.Visible = False

    StoreMemVars

    Call ResizeColumnHeader(lstExecutive, "60,32")

    Call FillCombo("SELECT DISTINCT TEAMNAME  from SMIS_vw_Srep WHERE LEN(TEAMNAME)>0", -1, 0, Combo1)


End Sub

Sub initMemvars()
    txtname.Text = ""
    txtFirstName.Text = ""
    txtLastName.Text = ""
    txtMiddleName.Text = ""

End Sub

Sub StoreMemVars()
    If Not rsSREP.EOF And Not rsSREP.BOF Then
        LabID.Caption = rsSREP!ID
        txtname.Text = Null2String(rsSREP!Name)
        txtFirstName.Text = Null2String(rsSREP!FIRSTNAME)
        txtLastName.Text = Null2String(rsSREP!lastname)
        txtMiddleName.Text = Null2String(rsSREP!Middlename)
        Combo1.Text = Null2String(rsSREP!TeamName)
        labSAECODE = Null2String(rsSREP!SAECODE)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set rsSREP = New ADODB.Recordset
    Dim sql                                                           As String

    sql = "SELECT  Lastname + ISNULL(',' + Firstname, '') + LEFT('.' + Middlename, 2) AS name, "
    sql = sql & " EMPLEVEL + EmpNo AS SAECODE, Firstname,Middlename,Lastname,"
    sql = sql & " (Select TEAMNAME FROM SMIS_SalesTeam WHERE SAECODE=EMPLEVEL + EmpNo ) as TEAMNAME,ID"
    sql = sql & " From HRMS_EmpInfo"
    sql = sql & " Where (IS_SAE = 1)  order by ID asc"

    rsSREP.Open sql, gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Private Sub lstExecutive_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsSREP.MoveFirst
    rsSREP.Find ("ID=" & lstExecutive.SelectedItem.SubItems(2))
    StoreMemVars
End Sub

Private Sub lstExecutive_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstExecutive
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

Private Sub lstExecutive_DblClick()

    If Not lstExecutive.ListItems.Count = 0 Then
        cmdEdit.Value = True
    End If
End Sub

Private Sub txtsearch_Change()
    If Trim(txtSearch.Text) = "" Then FillGrid Else FillSearchGrid (txtSearch.Text)
End Sub


'==========================================================================================
'FUNCTION / FEATURE :   FillSearchGrid:FillSearchGrid: ADDED FIX FOR SAE
'                       That are not visible FROM HRMS as To Declare SAE TEAM
'DATE STARTED       :6/5/200723:46
'LAST UPDATED       :6/5/200723:46
'DATABASE UPDATES   :
'WHO UPDATED        :AXP  6/5/2007
'UDPATING CODE    :AXP-6520071146
'==========================================================================================

Sub FillGrid()
    Dim rsSREP                                                        As ADODB.Recordset
    lstExecutive.Enabled = False
    lstExecutive.Sorted = False: lstExecutive.ListItems.Clear
    Set rsSREP = New ADODB.Recordset

    Dim sql                                                           As String
    sql = "SELECT  Lastname + ISNULL(',' + Firstname, '') + LEFT('.' + Middlename, 2) AS name, "
    sql = sql & " (Select TEAMNAME FROM SMIS_SalesTeam WHERE SAECODE=EMPLEVEL + EmpNo ) as TEAMNAME,ID"
    sql = sql & " From HRMS_EmpInfo"
    sql = sql & " Where (IS_SAE = 1)  order by name asc"
    Set rsSREP = gconDMIS.Execute(sql)
    If Not (rsSREP.EOF And rsSREP.BOF) Then
        Listview_Loadval Me.lstExecutive.ListItems, rsSREP
        lstExecutive.Refresh
        lstExecutive.Enabled = True
    End If

End Sub


Sub FillSearchGrid(xxx As String)
    Dim rsSREP                                                        As ADODB.Recordset

    lstExecutive.Sorted = False: lstExecutive.ListItems.Clear
    lstExecutive.Enabled = True
    Set rsSREP = New ADODB.Recordset
    Dim sql                                                           As String
    sql = "SELECT  Lastname + ISNULL(',' + Firstname, '') + LEFT('.' + Middlename, 2) AS name, "
    sql = sql & " (Select TEAMNAME FROM SMIS_SalesTeam WHERE SAECODE=EMPLEVEL + EmpNo ) as TEAMNAME,ID"
    sql = sql & " From HRMS_EmpInfo"
    sql = sql & " Where (IS_SAE = 1) AND LASTNAME like'" & ReplaceQuote(xxx) & "%' order by name asc"
    Set rsSREP = gconDMIS.Execute(sql)

    If Not (rsSREP.EOF And rsSREP.BOF) Then
        Listview_Loadval Me.lstExecutive.ListItems, rsSREP
        lstExecutive.Refresh
        lstExecutive.Enabled = True
    End If
End Sub
'Sub FillGrid()
'    Dim rsSREP                         As adodb.Recordset
'    lstExecutive.Sorted = False: lstExecutive.ListItems.Clear
'    Set rsSREP = New adodb.Recordset
'    Set rsSREP = gconDMIS.Execute("select Name,TeamName,ID from SMIS_vw_SRep order by name asc")
'    If Not (rsSREP.EOF And rsSREP.BOF) Then
'        Listview_Loadval Me.lstExecutive.ListItems, rsSREP
'        lstExecutive.Refresh
'    End If
'End Sub
'
'Sub FillSearchGrid(xxx As String)
'    Dim rsSREP                         As adodb.Recordset
'    lstExecutive.Sorted = False: lstExecutive.ListItems.Clear
'    Set rsSREP = New adodb.Recordset
'    Set rsSREP = gconDMIS.Execute("select Name,TeamName,ID from SMIS_vw_SRep WHERE name like'" & ReplaceQuote(xxx) & "%' order by name asc")
'    If Not (rsSREP.EOF And rsSREP.BOF) Then
'        Listview_Loadval Me.lstExecutive.ListItems, rsSREP
'        lstExecutive.Refresh
'    End If
'End Sub


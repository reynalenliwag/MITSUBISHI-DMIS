VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMSContractor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contractor Data Entry"
   ClientHeight    =   3975
   ClientLeft      =   720
   ClientTop       =   330
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "CSMS_Contractor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3975
   ScaleWidth      =   8910
   Begin VB.PictureBox Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3045
      Left            =   2670
      ScaleHeight     =   3015
      ScaleWidth      =   6135
      TabIndex        =   16
      Top             =   30
      Width           =   6165
      Begin VB.TextBox txtLastName 
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
         Left            =   1350
         TabIndex        =   22
         Top             =   1590
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.TextBox txtFirstName 
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
         Left            =   1350
         TabIndex        =   21
         Top             =   1980
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.TextBox txtMname 
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
         Left            =   1350
         TabIndex        =   20
         Top             =   2370
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1710
         TabIndex        =   19
         Top             =   1140
         Width           =   4125
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1710
         MaxLength       =   7
         TabIndex        =   18
         Top             =   330
         Width           =   1635
      End
      Begin VB.TextBox txtCompanyName 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1710
         TabIndex        =   17
         Top             =   720
         Width           =   4125
      End
      Begin Crystal.CrystalReport rptSA 
         Left            =   5460
         Top             =   2520
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Service Advisor's Master List"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Label labid 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   5430
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
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
         Height          =   210
         Left            =   270
         TabIndex        =   29
         Top             =   1650
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
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
         Height          =   210
         Left            =   270
         TabIndex        =   28
         Top             =   2010
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Initial"
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
         Height          =   210
         Left            =   90
         TabIndex        =   27
         Top             =   2430
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Height          =   210
         Left            =   870
         TabIndex        =   26
         Top             =   1260
         Width           =   765
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Height          =   210
         Left            =   1140
         TabIndex        =   25
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Contractor Name"
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
         Height          =   210
         Left            =   30
         TabIndex        =   24
         Top             =   840
         Width           =   1605
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   6135
         _Version        =   655364
         _ExtentX        =   10821
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "INFORMATION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16711680
         GradientColorDark=   8388608
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   2880
      ScaleHeight     =   945
      ScaleWidth      =   6315
      TabIndex        =   12
      Top             =   3060
      Width           =   6315
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   5220
         MouseIcon       =   "CSMS_Contractor.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_Contractor.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Exit Window"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   795
         Left            =   4500
         MouseIcon       =   "CSMS_Contractor.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_Contractor.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print this Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   3780
         MouseIcon       =   "CSMS_Contractor.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_Contractor.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Delete Selected Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   3060
         MouseIcon       =   "CSMS_Contractor.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_Contractor.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   2340
         MouseIcon       =   "CSMS_Contractor.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_Contractor.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Add Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   1620
         MouseIcon       =   "CSMS_Contractor.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_Contractor.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Find a Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   900
         MouseIcon       =   "CSMS_Contractor.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_Contractor.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Move to Next Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   180
         MouseIcon       =   "CSMS_Contractor.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_Contractor.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   7305
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   13
      Top             =   3045
      Width           =   1800
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   780
         MouseIcon       =   "CSMS_Contractor.frx":2D71
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_Contractor.frx":2EC3
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Cancel"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   60
         MouseIcon       =   "CSMS_Contractor.frx":3201
         MousePointer    =   99  'Custom
         Picture         =   "CSMS_Contractor.frx":3353
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.Frame fraDetails 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   60
      TabIndex        =   9
      Top             =   -45
      Width           =   2565
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
         Left            =   90
         MaxLength       =   35
         TabIndex        =   10
         Top             =   180
         Width           =   2415
      End
      Begin MSComctlLib.ListView lstContractor 
         Height          =   2805
         Left            =   90
         TabIndex        =   11
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   4948
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
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "CSMS_Contractor.frx":36A3
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Contractor"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   600
      TabIndex        =   8
      Top             =   270
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmCSMSContractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsContractor                                       As ADODB.Recordset
Dim AddorEdit                                          As String

Sub initMemvars()
    txtcode.Text = ""
    txtLastName.Text = ""
    txtFirstName.Text = ""
    txtMname.Text = ""
    txtCompanyName.Text = ""
    txtAddress.Text = ""
End Sub

Sub StoreMemVars()
    On Error Resume Next
    If Not rsContractor.EOF And Not rsContractor.BOF Then
        labid.Caption = rsContractor!ID
        txtcode.Text = Null2String(rsContractor!code)
        txtLastName.Text = Null2String(rsContractor!lastname)
        txtFirstName.Text = Null2String(rsContractor!Firstname)
        txtMname.Text = Null2String(rsContractor!MNAME)
        txtCompanyName.Text = Null2String(rsContractor!CompanyName)
        txtAddress.Text = Null2String(rsContractor!Address)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set rsContractor = New ADODB.Recordset
    rsContractor.Open "select * from CSMS_Contractor order by CompanyName asc", gconDMIS, adOpenKeyset
End Sub

Sub FillGrid()
    Dim rsServiceAdvisor                               As ADODB.Recordset
    lstContractor.Enabled = False: lstContractor.Sorted = False: lstContractor.ListItems.Clear
    Set rsServiceAdvisor = New ADODB.Recordset
    Set rsServiceAdvisor = gconDMIS.Execute("select CompanyName,Address from CSMS_Contractor order by CompanyName asc")
    If Not (rsServiceAdvisor.EOF And rsServiceAdvisor.BOF) Then
        Listview_Loadval Me.lstContractor.ListItems, rsServiceAdvisor
        lstContractor.Refresh
        lstContractor.Enabled = True
    End If
    Set rsServiceAdvisor = Nothing
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsServiceAdvisor                               As ADODB.Recordset
    lstContractor.Sorted = False: lstContractor.ListItems.Clear: lstContractor.Enabled = False
    Set rsServiceAdvisor = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsServiceAdvisor = gconDMIS.Execute("select CompanyName, Address from CSMS_Contractor where CompanyName like'" & XXX & "%'")
    If Not (rsServiceAdvisor.EOF And rsServiceAdvisor.BOF) Then
        Listview_Loadval Me.lstContractor.ListItems, rsServiceAdvisor
        lstContractor.Refresh
        lstContractor.Enabled = True
    End If
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "CONTRACTOR") = False Then Exit Sub

    Screen.MousePointer = 11
    PrintSQLReport rptSA, CSMS_REPORT_PATH & "Contractor.rpt", "", CSMS_REPORT_CONNECTION, 1
    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("V", "CONTRACTOR", "", labid, "", "CODE: " & txtcode, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    '    LogAudit "V", "CONTRACTOR INFORMATION REPORT "
    Screen.MousePointer = 0
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "CONTRACTOR") = False Then Exit Sub

    AddorEdit = "ADD"
    Frame1.Enabled = True
    fraDetails.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    On Error Resume Next
    txtcode.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    fraDetails.Enabled = True
    Picture1.Visible = True
    Picture2.Visible = False
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "CONTRACTOR") = False Then Exit Sub

    On Error GoTo ErrorCode

    If Not rsContractor.BOF Or Not rsContractor.EOF Then
        If MsgBox("Delete this Information", vbQuestion + vbYesNo, "Are you sure") = vbYes Then
            SQL_STATEMENT = "delete from CSMS_Contractor where id = " & labid.Caption
            gconDMIS.Execute SQL_STATEMENT
            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("X", "CONTRACTOR", SQL_STATEMENT, labid, "", "CODE: " & txtcode, "", "")
            'NEW LOG AUDIT-----------------------------------------------------

            SQL_STATEMENT = "DELETE FROM CSMS_CONTRACTORMONITORING WHERE CODE = '" & txtcode & "'"
            gconDMIS.Execute SQL_STATEMENT
            'LogAudit "X", "CONTRACTOR INFORMATION", "CODE/LASTNAME " & txtCode & "/" & txtLastName
            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("X", "CONTRACTOR", SQL_STATEMENT, labid, "", "CODE: " & txtcode, "", "")
            'NEW LOG AUDIT-----------------------------------------------------

            FillGrid
            ShowDeletedMsg
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
    If Function_Access(LOGID, "Acess_EDIT", "CONTRACTOR") = False Then Exit Sub

    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
    txtcode.SetFocus
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
    rsContractor.MoveNext
    If rsContractor.EOF Then
        rsContractor.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsContractor.MovePrevious
    If rsContractor.BOF Then
        rsContractor.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    'On Error GoTo Errorcode
    If IsNull(txtcode.Text) = True Then
        MsgSpeechBox "Contractor Code must not be empty"
        On Error Resume Next
        txtcode.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Dim rsfindDup                              As ADODB.Recordset
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select code from CSMS_Contractor where code = '" & txtcode.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "Contractor Code already exist!"
                On Error Resume Next
                txtcode.SetFocus
                Exit Sub
            End If
        End If
    End If
    If txtCompanyName.Text = "" Then                  'Or txtFirstName.Text = "" Then
        ShowIsRequiredMsg "Company Name cannot be Blank"
        On Error Resume Next
        txtCompanyName.SetFocus
        Exit Sub
    End If

    Dim VTXTCode, VTXTLASTNAME, VTXTFIRSTNAME          As String
    Dim VTXTMname, VTXTCompanyName, VTXTAddress        As String

    VTXTCode = N2Str2Null(txtcode.Text)
    VTXTLASTNAME = N2Str2Null(UCase(txtCompanyName.Text))
    VTXTFIRSTNAME = N2Str2Null(UCase(txtCompanyName.Text))
    VTXTMname = N2Str2Null(txtCompanyName.Text)
    VTXTCompanyName = N2Str2Null(txtCompanyName.Text)
    VTXTAddress = N2Str2Null(txtAddress.Text)

    If AddorEdit = "ADD" Then
        If Not rsContractor.EOF And Not rsContractor.BOF Then
            rsContractor.MoveLast
            labid.Caption = NumericVal(rsContractor!ID) + 1
        End If
        SQL_STATEMENT = "Insert into CSMS_Contractor" & _
                      " (code,lastname,firstname,Mname,CompanyName,Address)" & _
                      " values (" & VTXTCode & ", " & VTXTLASTNAME & ", " & VTXTFIRSTNAME & ", " & VTXTMname & ", " & _
                      " " & VTXTCompanyName & ", " & VTXTAddress & ")"
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("A", "CONTRACTOR", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtcode), "CODE", "CSMS_Contractor"), "", "CODE: " & txtcode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        gconDMIS.Execute ("INSERT INTO CSMS_CONTRACTORMONITORING VALUES('" & txtcode.Text & _
                          "',NULL,'" & txtCompanyName & "','AVAILABLE')")

        'LogAudit "A", "SERVICE ADVISOR INFORMATION", "CODE/LASTNAME " & txtCode & "/" & txtLastName
        ShowSuccessFullyAdded
    Else
        gconDMIS.Execute "update CSMS_Contractor set" & _
                       " code = " & VTXTCode & "," & _
                       " lastname = " & N2Str2Null(txtCompanyName) & "," & _
                       " firstname = " & N2Str2Null(txtCompanyName) & "," & _
                       " Mname = " & N2Str2Null(txtCompanyName) & "," & _
                       " CompanyName = " & VTXTCompanyName & "," & _
                       " Address = " & VTXTAddress & _
                       " where id = " & labid.Caption

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "CONTRACTOR", SQL_STATEMENT, labid, "", "CODE: " & txtcode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        'LogAudit "E", "CONTRACTOR INFORMATION", "CODE/LASTNAME " & txtCode & "/" & txtLastName
        ShowSuccessFullyUpdated
    End If

    FillGrid
    rsRefresh
    On Error Resume Next
    rsContractor.Find "id =" & labid.Caption
    cmdCancel.Value = True
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            If Picture1.Visible = True Then
                Unload frmALL_AuditInquiry

                frmALL_AuditInquiry.Show
                frmALL_AuditInquiry.ZOrder 0
                frmALL_AuditInquiry.Caption = "Audit Inquiry (TRANSACTIONS FOR FOLLOW UP)"
                Call frmALL_AuditInquiry.DisplayHistory(labid, "CONTRACTOR", "")
            End If
    End Select

End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    Frame1.Enabled = False
    textSearch.Text = "":

    FillGrid
    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub lstContractor_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    On Error Resume Next
    rsContractor.Bookmark = rsFind(rsContractor.Clone, "CompanyName", lstContractor.SelectedItem).Bookmark
    StoreMemVars
End Sub

Private Sub lstContractor_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstContractor
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

Private Sub lstContractor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (textSearch.Text)
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstContractor.Enabled = True Then
            lstContractor.SetFocus
        End If
    End If
End Sub


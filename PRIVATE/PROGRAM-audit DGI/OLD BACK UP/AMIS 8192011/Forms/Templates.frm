VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A06473E6-73D7-426E-82F2-6CD4F1FA4DBE}#1.0#0"; "wizMACBut.ocx"
Begin VB.Form frmAMISMASTERFILESTemplates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Entries Templates"
   ClientHeight    =   5580
   ClientLeft      =   1665
   ClientTop       =   1275
   ClientWidth     =   5745
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Templates.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   5745
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   -30
      ScaleHeight     =   855
      ScaleWidth      =   5760
      TabIndex        =   19
      Top             =   4635
      Width           =   5760
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
         Left            =   4965
         MouseIcon       =   "Templates.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "Templates.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   27
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
         Left            =   4275
         MouseIcon       =   "Templates.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "Templates.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Left            =   3585
         MouseIcon       =   "Templates.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "Templates.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Left            =   2895
         MouseIcon       =   "Templates.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "Templates.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   2205
         MouseIcon       =   "Templates.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "Templates.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Left            =   1515
         MouseIcon       =   "Templates.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "Templates.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   825
         MouseIcon       =   "Templates.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "Templates.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Left            =   135
         MouseIcon       =   "Templates.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "Templates.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   5625
      Begin VB.ComboBox cboJType 
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
         ForeColor       =   &H00973640&
         Height          =   330
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   4245
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
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   150
         Width           =   4245
      End
      Begin Crystal.CrystalReport rptTemplate_Header 
         Left            =   5100
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Account Headers"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Journal"
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
         TabIndex        =   5
         Top             =   570
         Width           =   855
      End
      Begin VB.Label Label3 
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
         Left            =   -60
         TabIndex        =   1
         Top             =   180
         Width           =   1245
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
         Left            =   3870
         TabIndex        =   3
         Top             =   180
         Width           =   465
      End
      Begin VB.Label labID 
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
         Left            =   4350
         TabIndex        =   4
         Top             =   210
         Width           =   225
      End
   End
   Begin VB.PictureBox picTemplates 
      BackColor       =   &H00EBFAFA&
      Height          =   1365
      Left            =   180
      ScaleHeight     =   1305
      ScaleWidth      =   5355
      TabIndex        =   11
      Top             =   2190
      Width           =   5415
      Begin VB.TextBox txtAccountCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00EBFAFA&
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
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   60
         Width           =   2145
      End
      Begin VB.ComboBox cboDescription 
         BackColor       =   &H00EFDFDF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   60
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   480
         Width           =   5175
      End
      Begin wizMacBut.MacBut cmdTempCancel 
         Height          =   345
         Left            =   3630
         TabIndex        =   18
         ToolTipText     =   "Cancel"
         Top             =   930
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   609
         Caption         =   "   Cancel"
      End
      Begin wizMacBut.MacBut cmdTempSave 
         Height          =   345
         Left            =   1890
         TabIndex        =   17
         ToolTipText     =   "Save Template"
         Top             =   930
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   609
         Caption         =   "    Save"
      End
      Begin wizMacBut.MacBut cmdTempDelete 
         Height          =   345
         Left            =   30
         TabIndex        =   16
         ToolTipText     =   "Delete Template"
         Top             =   930
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   609
         Caption         =   "   Delete"
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1770
         MaxLength       =   50
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   90
         Width           =   315
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Code :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   14
         Top             =   120
         Width           =   1545
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4230
      ScaleHeight     =   885
      ScaleWidth      =   1485
      TabIndex        =   28
      Top             =   4635
      Width           =   1485
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
         MouseIcon       =   "Templates.frx":2D71
         MousePointer    =   99  'Custom
         Picture         =   "Templates.frx":2EC3
         Style           =   1  'Graphical
         TabIndex        =   30
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
         MouseIcon       =   "Templates.frx":3201
         MousePointer    =   99  'Custom
         Picture         =   "Templates.frx":3353
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3735
      Left            =   60
      TabIndex        =   7
      Top             =   870
      Width           =   5625
      Begin VB.TextBox txtsearch 
         Height          =   375
         Left            =   1290
         TabIndex        =   32
         Top             =   150
         Width           =   4245
      End
      Begin MSComctlLib.ListView lstTemplates 
         Height          =   2835
         Left            =   60
         TabIndex        =   8
         Top             =   570
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   5001
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
         MouseIcon       =   "Templates.frx":36A3
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ACCOUNT NAME"
            Object.Width           =   9172
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Find Account"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   31
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00EBFAFA&
         BackStyle       =   0  'Transparent
         Caption         =   "Press <F3> to Add Entries for this Template"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   60
         TabIndex        =   9
         Top             =   3420
         Width           =   5505
      End
   End
   Begin wizButton.cmd cmdTemplates 
      Height          =   1485
      Left            =   120
      TabIndex        =   10
      Top             =   1500
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2619
      TX              =   "cmd1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "Templates.frx":3805
   End
End
Attribute VB_Name = "frmAMISMASTERFILESTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsChartAccount                                     As ADODB.Recordset
Dim rsTemplate_Header                                  As ADODB.Recordset
Dim AddorEdit, PrevCode                                As String
Attribute PrevCode.VB_VarUserMemId = 1073938434

Function SetAccCode(Acc As String) As String
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("select * from AMIS_ChartAccount where description = " & N2Str2Null(Acc))
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        SetAccCode = Null2String(rsChartAccount!ACCTCODE)
    End If
End Function

Sub InitDetails()
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("select Description from AMIS_ChartAccount order by AcctCode asc")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then Combo_Loadval cboDescription, rsChartAccount
    txtAccountCode.Text = ""
End Sub

Sub initMemvars()
    Frame1.Enabled = True
    txtDescription.Text = ""
End Sub

Sub rsRefresh()
    Set rsTemplate_Header = New ADODB.Recordset
    Set rsTemplate_Header = gconDMIS.Execute("select TemplateCode,description,JType from AMIS_Template_Header order by Description asc")
End Sub

Sub StoreEntry(XXX As Variant)
    Dim rsTemplate_Details                             As ADODB.Recordset
    Set rsTemplate_Details = New ADODB.Recordset
    Set rsTemplate_Details = gconDMIS.Execute("select * from AMIS_Template_Details where code = " & XXX)
    If Not rsTemplate_Details.EOF And Not rsTemplate_Details.BOF Then
        cmdTemplates.ZOrder 0: picTemplates.ZOrder 0
        txtCode.Text = rsTemplate_Details!Code
        txtAccountCode.Text = Null2String(rsTemplate_Details!AccountCode)
        cboDescription.Text = Null2String(rsTemplate_Details!DESCRIPTION)
    End If
End Sub

Sub StoreMemVars()
    If Not rsTemplate_Header.EOF And Not rsTemplate_Header.BOF Then
        Frame1.Enabled = False
        labID.Caption = rsTemplate_Header!TemplateCode
        txtDescription.Text = Null2String(rsTemplate_Header!DESCRIPTION)
        If Null2String(rsTemplate_Header!jtype) = "APJ" Then
            cboJTYPE.Text = "ACCOUNTS PAYABLE JOURNAL"
        ElseIf Null2String(rsTemplate_Header!jtype) = "CDJ" Then
            cboJTYPE.Text = "CASH DISBURSEMENT JOURNAL"
        ElseIf Null2String(rsTemplate_Header!jtype) = "SJ" Then
            cboJTYPE.Text = "ACCOUNTS RECEIVABLE JOURNAL"
        ElseIf Null2String(rsTemplate_Header!jtype) = "CRJ" Then
            cboJTYPE.Text = "CASH RECEIPTS JOURNAL"
        Else
            cboJTYPE.Text = "GENERAL JOURNAL"
        End If
        FillGrid
    Else
        MsgBox "No Such Record!"
        cmdAdd.Value = True
    End If
End Sub

Private Sub cboDescription_Click()
    txtAccountCode.Text = SetAccCode(cboDescription.Text)
End Sub

Private Sub cboDescription_LostFocus()
    txtAccountCode.Text = SetAccCode(cboDescription.Text)
End Sub

'Upating Code       : AXP-0713200713:54
Private Sub cmdAdd_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Add", "ACCOUNT ENTRIES TEMPLATES") = False Then Exit Sub

    AddorEdit = "ADD": initMemvars: Picture1.Visible = False: Picture2.Visible = True
    On Error Resume Next
    txtDescription.SetFocus
    lstTemplates.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False: Picture1.Visible = True: Picture2.Visible = False: StoreMemVars: fraDetails.Enabled = True: lstTemplates.Enabled = True
    lstTemplates.Enabled = True
End Sub

'Upating Code       : AXP-0713200713:55
Private Sub cmdDelete_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Delete", "ACCOUNT ENTRIES TEMPLATES") = False Then Exit Sub

    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from AMIS_Template_Header where TemplateCode = " & labID.Caption
        rsRefresh
        StoreMemVars
        NEW_LogAudit "X", "ACCOUNT ENTRIES TEMPLATE", SQL_STATEMENT, labID.Caption, "", txtAccountCode, "", ""
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:55
Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Edit", "ACCOUNT ENTRIES TEMPLATES") = False Then Exit Sub

    AddorEdit = "EDIT": Frame1.Enabled = True: Picture1.Visible = False: Picture2.Visible = True:
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim findStr                                        As String
    findStr = InputBox("Please Input Template Description ...", "Find")
    If findStr <> "" Then
        On Error GoTo ErrorCode
        rsTemplate_Header.Bookmark = rsFind(rsTemplate_Header.Clone, "Description", findStr).Bookmark
    End If
    StoreMemVars
    Exit Sub

ErrorCode:
    If Err.Number = 3021 Then
        MsgBox "Can't find " & findStr, vbOKOnly + vbExclamation, "Not Found"
        Resume Next
    End If
End Sub

'Upating Code       : AXP-0713200713:55
Private Sub cmdNext_Click()
    On Error GoTo ErrorCode:

    rsTemplate_Header.MoveNext
    If rsTemplate_Header.EOF Then
        rsTemplate_Header.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:55
Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode:

    rsTemplate_Header.MovePrevious
    If rsTemplate_Header.BOF Then
        rsTemplate_Header.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:55
Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Print", "ACCOUNT ENTRIES TEMPLATES") = False Then Exit Sub

    Screen.MousePointer = 11

    rptTemplate_Header.Reset
    rptTemplate_Header.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptTemplate_Header.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptTemplate_Header.ReportTitle = "Account Template"




    PrintSQLReport rptTemplate_Header, AMIS_REPORT_PATH & "AccountFiles\Template_Header.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    NEW_LogAudit "V", "ACCOUNT ENTRIES TEMPLATE", "", labID.Caption, "", txtAccountCode, "", ""
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:55
Private Sub cmdSave_Click()

    Dim vtxtDescription, VtxtJType                     As String
    On Error GoTo ErrorCode:

    vtxtDescription = N2Str2Null(txtDescription.Text)
    If cboJTYPE.Text = "ACCOUNTS PAYABLE JOURNAL" Then
        VtxtJType = "'APJ'"
    ElseIf cboJTYPE.Text = "CASH DISBURSEMENT JOURNAL" Then
        VtxtJType = "'CDJ'"
    ElseIf cboJTYPE.Text = "ACCOUNTS RECEIVABLE JOURNAL" Then
        VtxtJType = "'SJ'"
    ElseIf cboJTYPE.Text = "CASH RECEIPTS JOURNAL" Then
        VtxtJType = "'CRJ'"
    Else
        VtxtJType = "'GJ'"
    End If
    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "Insert into AMIS_Template_Header (Description,JType) values (" & vtxtDescription & "," & VtxtJType & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", "ACCOUNT ENTRIES TEMPLATE", SQL_STATEMENT, labID.Caption, "", txtAccountCode, "", ""
    Else
        SQL_STATEMENT = "update AMIS_Template_Header set Description = " & vtxtDescription & ", Jtype = " & VtxtJType & " where TemplateCode = " & labID
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "ACCOUNT ENTRIES TEMPLATE", SQL_STATEMENT, labID.Caption, "", txtAccountCode, "", ""
    End If
    rsRefresh
    FillGrid
    On Error Resume Next
    rsTemplate_Header.Find "Description = " & vtxtDescription
    cmdCancel.Value = True
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdTempCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdTemplates.ZOrder 1: picTemplates.ZOrder 1
End Sub

'Upating Code       : AXP-0713200713:56
Private Sub cmdTempDelete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorCode:
    If ShowConfirmDelete = True Then
        SQL_STATEMENT = "delete from AMIS_Template_Details Where Code = " & txtCode.Text
        gconDMIS.Execute SQL_STATEMENT
        rsRefresh
        StoreMemVars
        txtAccountCode = ""
        cboDescription = ""
        AddorEdit = ""
        NEW_LogAudit "XX", "ACCOUNT ENTRIES TEMPLATE", SQL_STATEMENT, labID.Caption, "", txtCode.Text, "", ""
        cmdTemplates.ZOrder 1: picTemplates.ZOrder 1

    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:56
Private Sub cmdTempSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrorCode:

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "Insert into AMIS_Template_Details " & _
                        "(TemplateCode,AccountCode,Description)" & _
                        " values (" & labID.Caption & "," & N2Str2Null(txtAccountCode.Text) & "," & N2Str2Null(UCase(cboDescription.Text)) & ")"

        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "AA", "ACCOUNT ENTRIES TEMPLATE", SQL_STATEMENT, labID.Caption, "", txtCode.Text, "", ""

    Else
        SQL_STATEMENT = "update AMIS_Template_Details Set " & _
                        " AccountCode = " & N2Str2Null(txtAccountCode.Text) & "," & _
                        " Description = " & N2Str2Null(UCase(cboDescription.Text)) & _
                        " Where Code = " & txtCode.Text
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "EE", "ACCOUNT ENTRIES TEMPLATE", SQL_STATEMENT, labID.Caption, "", txtCode.Text, "", ""
    End If
    rsRefresh
    rsTemplate_Header.Find "TemplateCode = " & labID.Caption
    StoreMemVars
    cmdTemplates.ZOrder 1: picTemplates.ZOrder 1
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub FillGrid()
    If labID = "" Then
        Exit Sub
    End If
    Dim rsTemplate_Details                             As ADODB.Recordset
    lstTemplates.Enabled = False
    lstTemplates.Sorted = False: lstTemplates.ListItems.Clear
    Set rsTemplate_Details = New ADODB.Recordset
    Set rsTemplate_Details = gconDMIS.Execute("select description,Code from AMIS_Template_Details WHERE TEMPLATECODE = " & labID)
    If Not (rsTemplate_Details.EOF And rsTemplate_Details.BOF) Then
        lstTemplates.Enabled = True
        Listview_Loadval Me.lstTemplates.ListItems, rsTemplate_Details
        lstTemplates.Refresh
        lstTemplates.Enabled = True
    Else
        lstTemplates.Enabled = False
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF1 And Shift = 1:
        If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
        Unload frmALL_AuditInquiry
        frmALL_AuditInquiry.Show
        frmALL_AuditInquiry.ZOrder 0
        frmALL_AuditInquiry.Caption = "ACCOUNT ENTRIES TEMPLATE"
        Call frmALL_AuditInquiry.DisplayHistory(labID, "ACCOUNT ENTRIES TEMPLATE")
    Case vbKeyF3
        AddorEdit = "ADD": cmdTemplates.ZOrder 0: picTemplates.ZOrder 0
        InitDetails
        On Error Resume Next
        cboDescription.SetFocus
    Case Else
        MoveKeyPress KeyCode


    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    cboJTYPE.Clear
    cboJTYPE.AddItem "ACCOUNTS PAYABLE JOURNAL"
    cboJTYPE.AddItem "CASH DISBURSEMENT JOURNAL"
    cboJTYPE.AddItem "ACCOUNTS RECEIVABLE JOURNAL"
    cboJTYPE.AddItem "CASH RECEIPTS JOURNAL"
    cboJTYPE.AddItem "GENERAL JOURNAL"
initMemvars:     rsRefresh: StoreMemVars: cmdTemplates.ZOrder 1: picTemplates.ZOrder 1: InitDetails
    Screen.MousePointer = 0
End Sub

Private Sub lstTemplates_DblClick()
    On Error Resume Next
    StoreEntry lstTemplates.SelectedItem.SubItems(1)
    AddorEdit = "EDIT"
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub
Sub Fillsearch()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim nard                                           As ListItem
    Dim keyword                                        As String
    keyword = Trim(txtSearch.Text)

    SQL = "SELECT * from AMIS_Template_Header where description like '" & keyword & "%'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)
    lstTemplates.ListItems.Clear
    Do While Not RS.EOF
        Set nard = lstTemplates.ListItems.Add(, , RS!DESCRIPTION)
        nard.SubItems(1) = Null2String(RS!TemplateCode)
        RS.MoveNext
    Loop
    lstTemplates.Enabled = True
    Set RS = Nothing
End Sub

Private Sub txtSearch_Change()
    Fillsearch
End Sub

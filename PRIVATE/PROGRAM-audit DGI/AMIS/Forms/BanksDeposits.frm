VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmAMISMASTERFILEBanks2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Banks Master List"
   ClientHeight    =   5040
   ClientLeft      =   990
   ClientTop       =   855
   ClientWidth     =   5760
   ForeColor       =   &H00FFFFFF&
   Icon            =   "BanksDeposits.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   5760
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   45
      ScaleHeight     =   855
      ScaleWidth      =   5625
      TabIndex        =   11
      Top             =   4125
      Width           =   5625
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
         MouseIcon       =   "BanksDeposits.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "BanksDeposits.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   19
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
         MouseIcon       =   "BanksDeposits.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "BanksDeposits.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   18
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
         MouseIcon       =   "BanksDeposits.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "BanksDeposits.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   17
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
         MouseIcon       =   "BanksDeposits.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "BanksDeposits.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   16
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
         MouseIcon       =   "BanksDeposits.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "BanksDeposits.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   15
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
         MouseIcon       =   "BanksDeposits.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "BanksDeposits.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   14
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
         MouseIcon       =   "BanksDeposits.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "BanksDeposits.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   13
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
         MouseIcon       =   "BanksDeposits.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "BanksDeposits.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4185
      ScaleHeight     =   885
      ScaleWidth      =   1485
      TabIndex        =   20
      Top             =   4125
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
         MouseIcon       =   "BanksDeposits.frx":2D71
         MousePointer    =   99  'Custom
         Picture         =   "BanksDeposits.frx":2EC3
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
         Left            =   30
         MouseIcon       =   "BanksDeposits.frx":3201
         MousePointer    =   99  'Custom
         Picture         =   "BanksDeposits.frx":3353
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin Crystal.CrystalReport rptBanks 
      Left            =   2340
      Top             =   150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Banks"
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
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   5625
      Begin VB.ComboBox cboAcctCode 
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
         Left            =   2490
         TabIndex        =   3
         Text            =   "cboAcctCode"
         Top             =   1350
         Width           =   3045
      End
      Begin VB.TextBox txtBankAcctNo 
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
         Left            =   150
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1320
         Width           =   2235
      End
      Begin VB.TextBox txtBankName 
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
         Left            =   1260
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   540
         Width           =   4245
      End
      Begin VB.TextBox txtBankCode 
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
         Left            =   1260
         MaxLength       =   8
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Chart of Account Code"
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
         Height          =   240
         Left            =   2520
         TabIndex        =   24
         Top             =   1050
         Width           =   2430
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Acct No."
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
         Height          =   240
         Left            =   135
         TabIndex        =   23
         Top             =   1020
         Width           =   1410
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
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
         Height          =   240
         Left            =   135
         TabIndex        =   6
         Top             =   600
         Width           =   1080
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
         Left            =   4710
         TabIndex        =   8
         Top             =   600
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
         TabIndex        =   7
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Code"
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
         Height          =   240
         Left            =   135
         TabIndex        =   5
         Top             =   210
         Width           =   1035
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   2325
      Left            =   30
      TabIndex        =   9
      Top             =   1770
      Width           =   5625
      Begin MSComctlLib.ListView lstBanks 
         Height          =   2115
         Left            =   30
         TabIndex        =   10
         Top             =   150
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   3731
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
         MouseIcon       =   "BanksDeposits.frx":36A3
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "BANK CODE"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "BANK NAME"
            Object.Width           =   6702
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAMISMASTERFILEBanks2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBanks                                                 As ADODB.Recordset
Dim rsCOA                                                   As ADODB.Recordset
Dim AddorEdit                                               As String

Sub initMemvars()
    Frame1.Enabled = True
    txtBankCode.Text = vbNullString
    txtBankName.Text = vbNullString
    txtBankAcctNo.Text = vbNullString
    InitCbo
End Sub

Sub InitCbo()
    Set rsCOA = New ADODB.Recordset
    If COMPANY_CODE = "DGI" Then
        Set rsCOA = gconDMIS.Execute("Select AcctCode,Description from AMIS_ChartAccount Where Titles = '1101' OR LEFT(DESCRIPTION,11) ='UNPRESENTED' Order by AcctCode asc")
    Else
        Set rsCOA = gconDMIS.Execute("Select AcctCode,Description from AMIS_ChartAccount Where Titles = '1101' Order by AcctCode asc")
    End If
    If Not rsCOA.EOF And Not rsCOA.BOF Then
        rsCOA.MoveFirst: cboAcctCode.Clear
        Do While Not rsCOA.EOF
            cboAcctCode.AddItem Null2String(rsCOA!AcctCode)    '+ " " + Null2String(rsCOA!Description)
            rsCOA.MoveNext
        Loop
    End If
    Set rsCOA = Nothing
End Sub

Sub rsRefresh()
    Set rsBanks = New ADODB.Recordset
    rsBanks.Open "select * from ALL_BankDeposits order by BankCode asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    FillGrid
End Sub

Sub StoreEntry(XXX As Variant)
    Dim rsBanks2                                            As ADODB.Recordset
    Set rsBanks2 = New ADODB.Recordset
    rsBanks2.Open "select * from ALL_BankDeposits where ID = " & XXX, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsBanks2.EOF And Not rsBanks2.BOF Then
        fraDetails.Enabled = False
        lstBanks.Enabled = False
        labID.Caption = rsBanks2!ID
        txtBankCode.Text = Null2String(rsBanks2!bankcode)
        txtBankName.Text = Null2String(rsBanks2!BankName)
    End If
End Sub

Sub StoreMemVars()
    If Not rsBanks.EOF And Not rsBanks.BOF Then
        Frame1.Enabled = False
        labID.Caption = rsBanks!ID
        txtBankCode.Text = Null2String(rsBanks!bankcode)
        txtBankName.Text = Null2String(rsBanks!BankName)
        txtBankAcctNo.Text = Null2String(rsBanks!BankAcctNo)
        cboAcctCode.Text = Null2String(rsBanks!AcctCode)
    Else
        lstBanks.ListItems.Clear
        MsgBox "No Such Record!"
        cmdAdd.Value = True
    End If
End Sub

'Upating Code       : AXP-0713200713:53
Private Sub cmdAdd_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Add", "BANKS") = False Then Exit Sub

    AddorEdit = "ADD"
    initMemvars
    Picture1.Visible = False
    Picture2.Visible = True
    lstBanks.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstBanks.Enabled = True
    fraDetails.Enabled = True
    StoreMemVars
    FillGrid
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "BANKS") = False Then Exit Sub
    Dim lngCount                                            As Long


    lngCount = gconDMIS.Execute("SELECT COUNT(*) FROM CMIS_BANKDEPO WHERE DEPOSIT_TO=" & N2Str2Null(txtBankCode)).Fields(0).Value
    If lngCount > 0 Then
        MsgBox "Bank Record Exists in Bank Deposit(s)." & vbCrLf & "Cannot delete Record.", vbInformation
        Exit Sub
    End If

    If MsgBox("Delete Current Record", vbQuestion + vbYesNo, "Delete") = vbYes Then
        SQL_STATEMENT = "delete from ALL_BankDeposits where ID = " & lstBanks.SelectedItem.SubItems(2)
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "X", "BANK MASTER FILE", SQL_STATEMENT, labID.Caption, "", txtBankCode, "", ""
    End If
    rsRefresh
    StoreMemVars
End Sub

'Upating Code       : AXP-0713200713:53
Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Edit", "BANKS") = False Then Exit Sub

    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    StoreEntry (lstBanks.SelectedItem.SubItems(2))
    lstBanks.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200713:52
Private Sub cmdFind_Click()
    Dim findStr                                             As String
    findStr = InputBox("Please Input Banks ...", "Find")
    If findStr <> "" Then
        On Error Resume Next
        rsBanks.Bookmark = rsFind(rsBanks.Clone, "BankCode", findStr).Bookmark
        If Err.Number = 3021 Then
            On Error GoTo ErrorCode
            rsBanks.Bookmark = rsFind(rsBanks.Clone, "BankName", findStr).Bookmark
        End If
    End If
    StoreMemVars
    Exit Sub

ErrorCode:
    If Err.Number = 3021 Then
        MsgBox "Can't find " & findStr, vbOKOnly + vbExclamation, "Not Found"
        Resume Next
    End If

End Sub

'Upating Code       : AXP-0713200713:52
Private Sub cmdNext_Click()
    On Error GoTo ErrorCode:

    rsBanks.MoveNext
    If rsBanks.EOF Then
        rsBanks.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:52
Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode:

    rsBanks.MovePrevious
    If rsBanks.BOF Then
        rsBanks.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:53
Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Print", "BANKS") = False Then Exit Sub

    Screen.MousePointer = 11
    rptBanks.Reset
    rptBanks.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptBanks.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptBanks.ReportTitle = " BANKS "
    PrintSQLReport rptBanks, AMIS_REPORT_PATH & "\Files\Banks.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    NEW_LogAudit "V", "BANK MASTER FILE", "", labID.Caption, "", txtBankCode, "", ""
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:53
Private Sub cmdSave_Click()
'    Dim CMD                                            As ADODB.Command
'    Dim PARAM                                          As ADODB.Parameter
'    Set CMD = New ADODB.Command
'    CMD.CommandText = "sp_manage_Banks"
'    CMD.CommandType = adCmdStoredProc
'    CMD.ActiveConnection = gconDMIS
'    MsgBox labID
'    CMD.Parameters.Append CMD.CreateParameter("@ID", adInteger, adParamInput, 0, labID)
'    CMD.Parameters.Append CMD.CreateParameter("@BANKCODE", adVarChar, adParamInput, 6, txtBankCode)
'    CMD.Parameters.Append CMD.CreateParameter("@BANKNAME", adVarChar, adParamInput, 50, txtBankName)
'    CMD.Execute
'    rsRefresh
'    rsBanks.Find ("id=" & labID)
'    StoreMemvars
'    cmdCancel.Value = True
'    Exit Sub



    Dim VtxtBankCode                                        As String
    Dim VtxtBankName                                        As String
    Dim VtxtBankAcNo                                        As String
    Dim VcboAcctCode                                        As String

    'On Error GoTo Errorcode:

    VtxtBankCode = N2Str2Null(txtBankCode.Text)
    VtxtBankName = N2Str2Null(txtBankName.Text)
    VtxtBankAcNo = N2Str2Null(txtBankAcctNo.Text)
    VcboAcctCode = N2Str2Null(cboAcctCode.Text)

    If AddorEdit = "ADD" Then
        Dim rsBanksDup                                      As ADODB.Recordset
        Set rsBanksDup = New ADODB.Recordset
        rsBanksDup.Open "select BankCode from ALL_BankDeposits where BankCode = " & VtxtBankCode, gconDMIS
        If Not rsBanksDup.EOF And Not rsBanksDup.BOF Then
            MsgBox "Bank Code Already Exist!", vbCritical, "Duplicate Bank Code Not Allowed"
            Exit Sub
        End If

        SQL_STATEMENT = "Insert into ALL_BankDeposits " & _
                        "(BankCode,BankName,BankAcctNo,AcctCode) " & _
                        " values (" & VtxtBankCode & _
                        ", " & VtxtBankName & _
                        ", " & VtxtBankAcNo & _
                        ", " & VcboAcctCode & ")"

        gconDMIS.Execute SQL_STATEMENT

        NEW_LogAudit "A", "BANK MASTER FILE", SQL_STATEMENT, labID.Caption, "", txtBankCode, "", N2Str2Null(VtxtBankName)

    Else
        SQL_STATEMENT = "update ALL_BankDeposits set" & _
                        " BankCode = " & VtxtBankCode & ", " & _
                        " BankName = " & VtxtBankName & ", " & _
                        " BankAcctNo = " & VtxtBankAcNo & ", " & _
                        " AcctCode = " & VcboAcctCode & _
                        " where ID = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT


        NEW_LogAudit "E", "BANK MASTER FILE", SQL_STATEMENT, labID.Caption, "", txtBankCode, "", N2Str2Null(VtxtBankName)

    End If
    rsRefresh
    FillGrid
    On Error Resume Next
    rsBanks.Find "ID = " & labID.Caption
    cmdCancel.Value = True
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub FillGrid()
    Dim rsBanks2                                            As ADODB.Recordset
    lstBanks.Enabled = False
    lstBanks.Sorted = False: lstBanks.ListItems.Clear
    Set rsBanks2 = New ADODB.Recordset
    Set rsBanks2 = gconDMIS.Execute("select bankcode,bankname,ID from ALL_BankDeposits")
    If Not (rsBanks2.EOF And rsBanks2.BOF) Then
        lstBanks.Enabled = True
        Listview_Loadval Me.lstBanks.ListItems, rsBanks2
        lstBanks.Refresh
        lstBanks.Enabled = True
    Else
        lstBanks.Enabled = False
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode

    Select Case KeyCode
    Case vbKeyF1 And Shift = 1:

        If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
        Unload frmALL_AuditInquiry
        frmALL_AuditInquiry.Show
        frmALL_AuditInquiry.ZOrder 0
        frmALL_AuditInquiry.Caption = "BANK MASTER FILE"
        Call frmALL_AuditInquiry.DisplayHistory(labID, "BANK MASTER FILE")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    initMemvars
    StoreMemVars
    FillGrid
    Screen.MousePointer = 0
End Sub

Private Sub lstBanks_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstBanks
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

Private Sub lstBanks_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstBanks_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsBanks.Bookmark = rsFind(rsBanks.Clone, "bankcode", Item).Bookmark
    StoreMemVars
End Sub

Private Sub txtBankCode_LostFocus()
    txtBankCode.Text = UCase(txtBankCode.Text)
End Sub


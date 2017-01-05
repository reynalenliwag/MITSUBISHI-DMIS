VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRMSDeductionCodeMaterFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deduction Code Master File"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   Icon            =   "frmHRMSDeductionCodeMaterFile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4740
   ScaleWidth      =   7890
   Begin VB.PictureBox picDed 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1425
      Left            =   2040
      ScaleHeight     =   1425
      ScaleWidth      =   5865
      TabIndex        =   12
      Top             =   120
      Width           =   5865
      Begin VB.OptionButton OPtEntry 
         Caption         =   "Base on Salary"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3690
         TabIndex        =   23
         Top             =   1050
         Width           =   1935
      End
      Begin VB.OptionButton OPtEntry 
         Caption         =   "By Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2070
         TabIndex        =   22
         Top             =   1050
         Width           =   1455
      End
      Begin VB.OptionButton OPtEntry 
         Caption         =   "By Time"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   750
         TabIndex        =   21
         Top             =   1050
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox txtDesc 
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
         Height          =   360
         Left            =   750
         TabIndex        =   14
         Top             =   600
         Width           =   5025
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
         Height          =   360
         Left            =   750
         MaxLength       =   2
         TabIndex        =   13
         Top             =   180
         Width           =   1155
      End
      Begin Crystal.CrystalReport rptDeduction 
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
      Begin VB.Label lblID 
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
         Left            =   3900
         TabIndex        =   17
         Top             =   180
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   -30
         TabIndex        =   16
         Top             =   660
         Width           =   705
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
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   1980
      ScaleHeight     =   2295
      ScaleWidth      =   5865
      TabIndex        =   10
      Top             =   1350
      Width           =   5865
      Begin MSComctlLib.ListView lstDed 
         Height          =   1995
         Left            =   90
         TabIndex        =   11
         Top             =   210
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3519
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
         Appearance      =   0
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
         MouseIcon       =   "frmHRMSDeductionCodeMaterFile.frx":058A
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Deduction Description"
            Object.Width           =   6702
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   4560
      Left            =   60
      ScaleHeight     =   4500
      ScaleWidth      =   1845
      TabIndex        =   9
      Top             =   60
      Width           =   1905
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   6960
         Left            =   0
         Picture         =   "frmHRMSDeductionCodeMaterFile.frx":06EC
         Top             =   0
         Width           =   9915
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2220
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   0
      Top             =   3690
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
         MouseIcon       =   "frmHRMSDeductionCodeMaterFile.frx":14449
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSDeductionCodeMaterFile.frx":1459B
         Style           =   1  'Graphical
         TabIndex        =   1
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
         MouseIcon       =   "frmHRMSDeductionCodeMaterFile.frx":14901
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSDeductionCodeMaterFile.frx":14A53
         Style           =   1  'Graphical
         TabIndex        =   2
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
         MouseIcon       =   "frmHRMSDeductionCodeMaterFile.frx":14DB9
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSDeductionCodeMaterFile.frx":14F0B
         Style           =   1  'Graphical
         TabIndex        =   3
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
         MouseIcon       =   "frmHRMSDeductionCodeMaterFile.frx":15236
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSDeductionCodeMaterFile.frx":15388
         Style           =   1  'Graphical
         TabIndex        =   4
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
         MouseIcon       =   "frmHRMSDeductionCodeMaterFile.frx":156E4
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSDeductionCodeMaterFile.frx":15836
         Style           =   1  'Graphical
         TabIndex        =   5
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
         MouseIcon       =   "frmHRMSDeductionCodeMaterFile.frx":15B49
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSDeductionCodeMaterFile.frx":15C9B
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton k 
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
         MouseIcon       =   "frmHRMSDeductionCodeMaterFile.frx":15F95
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSDeductionCodeMaterFile.frx":160E7
         Style           =   1  'Graphical
         TabIndex        =   7
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
         MouseIcon       =   "frmHRMSDeductionCodeMaterFile.frx":1643F
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSDeductionCodeMaterFile.frx":16591
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   6360
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   18
      Top             =   3690
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
         MouseIcon       =   "frmHRMSDeductionCodeMaterFile.frx":168F0
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSDeductionCodeMaterFile.frx":16A42
         Style           =   1  'Graphical
         TabIndex        =   19
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
         MouseIcon       =   "frmHRMSDeductionCodeMaterFile.frx":16D80
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSDeductionCodeMaterFile.frx":16ED2
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmHRMSDeductionCodeMaterFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDED                                                             As ADODB.Recordset
Dim AddorEdit                                                         As String

Sub StoreMemVars()
    If Not (rsDED.EOF And rsDED.BOF) Then
        lblID.Caption = rsDED!ID
        txtCode.Text = Null2String(rsDED!CODE)
        txtDesc.Text = Null2String(rsDED!Description)
        If Null2String(rsDED!ENTRYBY) = "TIME" Then OPtEntry(0).Value = True
        If Null2String(rsDED!ENTRYBY) = "AMOUNT" Then OPtEntry(1).Value = True
        If Null2String(rsDED!ENTRYBY) = "SALARY" Then OPtEntry(2).Value = True
    Else
        ShowNoRecord
    End If
End Sub

Sub rsrefresh()
    Set rsDED = New ADODB.Recordset
    rsDED.Open "select * from HRMS_DeductionCode order by description ASC", gconDMIS, adOpenKeyset
End Sub

Sub InitMemvars()
    picDed.Enabled = True

    txtCode.Text = ""
    txtDesc.Text = ""
End Sub

Sub DisplayAllDeductionCode()
    Dim rstmp                                                         As New ADODB.Recordset
    Dim ITEM                                                          As ListItem

    Set rstmp = gconDMIS.Execute("Select * From HRMS_DeductionCode Order By Description asc")
    lstDed.ListItems.Clear
    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            Set ITEM = lstDed.ListItems.Add(, , Null2String(rstmp!CODE))
            ITEM.SubItems(1) = Null2String(rstmp!Description)
            ITEM.SubItems(2) = Null2String(rstmp!ENTRYBY)
            ITEM.SubItems(3) = Null2String(rstmp!ID)

            rstmp.MoveNext
        Loop
    End If

    Set rstmp = Nothing
End Sub

Private Sub cmdAdd_Click()

    If Function_Access(LOGID, "Acess_Add", "FILES DEDUCTION CODES") = False Then Exit Sub

    AddorEdit = "ADD"
    InitMemvars
    lstDed.Enabled = False

    Picture1.Visible = False
    Picture2.Visible = True

    txtCode.SetFocus
End Sub

Private Sub cmdCancel_Click()
    picDed.Enabled = False

    Picture1.Visible = True
    Picture2.Visible = False
    lstDed.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "ACESS_DELETE", "FILES DEDUCTION CODES") = False Then Exit Sub
    If Not txtCode.Text = "" Then
        If MsgBox("Delete This Deduction type", vbQuestion + vbYesNo, "Are You Sure") = vbYes Then
            SQL_STATEMENT = "Delete From HRMS_DeductionCode Where Code = '" & txtCode.Text & "'"
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "X", "FILE DEDUCTION CODE", SQL_STATEMENT, txtCode.Text, "", "", "", ""
            SQL_STATEMENT = ""
            ShowDeletedMsg
            DisplayAllDeductionCode
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "FILES DEDUCTION CODES") = False Then Exit Sub

    AddorEdit = "EDIT"
    picDed.Enabled = True

    Picture1.Visible = False
    Picture2.Visible = True
    lstDed.Enabled = False
End Sub

Private Sub cmdExit_Click()
    UnloadForm Me
End Sub

Private Sub cmdFind_Click()
    MsgBox "Use the List view to find...", vbInformation, "Find"
End Sub

Private Sub cmdNext_Click()
    rsDED.MoveNext
    If rsDED.EOF Then
        rsDED.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsDED.MovePrevious
    If rsDED.BOF Then
        rsDED.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "FILES DEDUCTION CODES") = False Then Exit Sub
    LogAudit "V", "PRINT DEDUCTION RECORD", ""

    Screen.MousePointer = 11
    rptDeduction.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
    rptDeduction.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
    rptDeduction.Formulas(2) = "COMPANYTIN = '" & COMPANY_TIN & "'"
    rptDeduction.Formulas(3) = "PRINTBY = '" & LOGNAME & "'"

    PrintSQLReport rptDeduction, HRMS_REPORT_PATH & "dEDUCTION List.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    'On Error GoTo Errorcode
    Dim DEDCODE                                                       As String
    Dim DEDDESC                                                       As String
    Dim DEDENTRY                                                      As String

    If txtCode.Text = "" Then MsgBox "Deduction Code cannot be Blank", vbInformation, "Incomplete Data": txtCode.SetFocus: Exit Sub
    If txtDesc.Text = "" Then MsgBox "Deduction Description Cannot be Blank", vbInformation, "Incomplete Data": txtDesc.SetFocus: Exit Sub

    DEDCODE = N2Str2Null(txtCode.Text)
    DEDDESC = N2Str2Null(txtDesc.Text)
    If OPtEntry(0).Value = True Then DEDENTRY = "TIME"
    If OPtEntry(1).Value = True Then DEDENTRY = "AMOUNT"
    If OPtEntry(2).Value = True Then DEDENTRY = "SALARY"

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "Insert into HRMS_DeductionCode" & _
                        "(code,Description,EntryBy) " & _
                      " values (" & DEDCODE & ", " & _
                        "" & DEDDESC & ",'" & DEDENTRY & "')"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", "FILE DEDUCTION CODE", SQL_STATEMENT, DEDCODE, "", "", "", ""
        SQL_STATEMENT = ""
        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update HRMS_DeductionCode set" & _
                      " description = " & DEDDESC & "," & _
                      " Code = " & DEDCODE & "," & _
                      " EntryBy = '" & DEDENTRY & _
                        "' where id = " & lblID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "FILE DEDUCTION CODE", SQL_STATEMENT, DEDCODE, "", "", "", ""
        SQL_STATEMENT = ""
        ShowSuccessFullyUpdated
    End If

    rsrefresh
    DisplayAllDeductionCode
    cmdCancel.Value = True
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (FILE DEDUCTION CODE)"
            Call frmALL_AuditInquiry.DisplayHistory(txtCode, "FILE DEDUCTION CODE")
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    DisplayAllDeductionCode
    rsrefresh
    StoreMemVars
End Sub

Private Sub lstDed_Click()
    Dim INDEX                                                         As Double

    If Not lstDed.ListItems.count = 0 Then
        With lstDed
            INDEX = .SelectedItem.INDEX

            txtCode.Text = .ListItems(INDEX).Text
            txtDesc.Text = .ListItems(INDEX).SubItems(1)

            If .ListItems(INDEX).SubItems(2) = "TIME" Then OPtEntry(0).Value = True
            If .ListItems(INDEX).SubItems(2) = "AMOUNT" Then OPtEntry(1).Value = True
            If .ListItems(INDEX).SubItems(2) = "SALARY" Then OPtEntry(2).Value = True

            lblID.Caption = .ListItems(INDEX).SubItems(3)


        End With
    End If
End Sub


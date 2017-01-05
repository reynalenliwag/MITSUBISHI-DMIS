VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRMS_LoanCodes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Code Master File"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHRMS_LoanCodes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   7995
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2310
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   9
      Top             =   3330
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
         MouseIcon       =   "frmHRMS_LoanCodes.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_LoanCodes.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   17
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
         MouseIcon       =   "frmHRMS_LoanCodes.frx":0A42
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_LoanCodes.frx":0B94
         Style           =   1  'Graphical
         TabIndex        =   16
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
         MouseIcon       =   "frmHRMS_LoanCodes.frx":0EFA
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_LoanCodes.frx":104C
         Style           =   1  'Graphical
         TabIndex        =   15
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
         MouseIcon       =   "frmHRMS_LoanCodes.frx":1377
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_LoanCodes.frx":14C9
         Style           =   1  'Graphical
         TabIndex        =   14
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
         MouseIcon       =   "frmHRMS_LoanCodes.frx":1825
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_LoanCodes.frx":1977
         Style           =   1  'Graphical
         TabIndex        =   13
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
         MouseIcon       =   "frmHRMS_LoanCodes.frx":1C8A
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_LoanCodes.frx":1DDC
         Style           =   1  'Graphical
         TabIndex        =   12
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
         MouseIcon       =   "frmHRMS_LoanCodes.frx":20D6
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_LoanCodes.frx":2228
         Style           =   1  'Graphical
         TabIndex        =   11
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
         MouseIcon       =   "frmHRMS_LoanCodes.frx":2580
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_LoanCodes.frx":26D2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   4140
      Left            =   120
      ScaleHeight     =   4080
      ScaleWidth      =   1845
      TabIndex        =   8
      Top             =   90
      Width           =   1905
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   6960
         Left            =   0
         Picture         =   "frmHRMS_LoanCodes.frx":2A31
         Top             =   -180
         Width           =   9915
      End
   End
   Begin VB.PictureBox picLoan 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1035
      Left            =   2070
      ScaleHeight     =   1035
      ScaleWidth      =   5865
      TabIndex        =   0
      Top             =   90
      Width           =   5865
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
         MaxLength       =   4
         TabIndex        =   2
         Top             =   180
         Width           =   1155
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
         TabIndex        =   1
         Top             =   600
         Width           =   5025
      End
      Begin Crystal.CrystalReport rptLoan 
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
         Left            =   0
         TabIndex        =   5
         Top             =   240
         Width           =   675
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
         TabIndex        =   4
         Top             =   660
         Width           =   705
      End
      Begin VB.Label labid 
         Caption         =   "ID"
         Height          =   285
         Left            =   3900
         TabIndex        =   3
         Top             =   180
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   2040
      ScaleHeight     =   2295
      ScaleWidth      =   5865
      TabIndex        =   6
      Top             =   1020
      Width           =   5865
      Begin MSComctlLib.ListView lstLoan 
         Height          =   2025
         Left            =   90
         TabIndex        =   7
         Top             =   120
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3572
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
         MouseIcon       =   "frmHRMS_LoanCodes.frx":1678E
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Loan Description"
            Object.Width           =   6702
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   6450
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   18
      Top             =   3330
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
         MouseIcon       =   "frmHRMS_LoanCodes.frx":168F0
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_LoanCodes.frx":16A42
         Style           =   1  'Graphical
         TabIndex        =   20
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
         MouseIcon       =   "frmHRMS_LoanCodes.frx":16D80
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_LoanCodes.frx":16ED2
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmHRMS_LoanCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSLOAN                                                            As ADODB.Recordset
Dim ADDOREDIT                                                         As String

Sub StoreMemVars()
    If Not (RSLOAN.EOF And RSLOAN.BOF) Then
        labid.Caption = RSLOAN!ID
        txtCode.Text = Null2String(RSLOAN!CODE)
        txtDesc.Text = Null2String(RSLOAN!Description)
    Else
        ShowNoRecord
    End If
End Sub

Sub rsrefresh()
    Set RSLOAN = New ADODB.Recordset
    RSLOAN.Open "select * from HRMS_LoanCode order by description ASC", gconDMIS, adOpenKeyset
End Sub

Sub InitMemvars()
    picLoan.Enabled = True

    txtCode.Text = ""
    txtDesc.Text = ""
End Sub

Sub DisplayAllLoanCode()
    Dim RSTMP                                                         As New ADODB.Recordset
    Dim ITEM                                                          As ListItem

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_LoanCode Order By Description asc")
    lstLoan.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lstLoan.ListItems.Add(, , Null2String(RSTMP!CODE))
            ITEM.SubItems(1) = Null2String(RSTMP!Description)
            ITEM.SubItems(2) = Null2String(RSTMP!ID)


            RSTMP.MoveNext
        Loop
    End If

    Set RSTMP = Nothing
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "FILES LOAD CODE") = False Then Exit Sub

    ADDOREDIT = "ADD"
    InitMemvars
    lstLoan.Enabled = False

    Picture1.Visible = False
    Picture2.Visible = True

    txtCode.SetFocus
End Sub

Private Sub cmdCancel_Click()
    picLoan.Enabled = False

    Picture1.Visible = True
    Picture2.Visible = False
    lstLoan.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "FILES LOAD CODE") = False Then Exit Sub

    If Not lstLoan.ListItems.count = 0 Then
        If MsgBox("Delete Loan Type", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            gconDMIS.Execute ("Delete From HRMS_LoanCode Where Code = '" & txtCode.Text & "'")

            LogAudit "X", "DELETE LOAN CODE", txtCode.Text
            ShowDeletedMsg
            DisplayAllLoanCode
        End If
    Else
        MsgBox "Theres nothing to Delete", vbInformation, "Information"
    End If
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "FILES LOAD CODE") = False Then Exit Sub

    ADDOREDIT = "EDIT"
    picLoan.Enabled = True

    Picture1.Visible = False
    Picture2.Visible = True
    lstLoan.Enabled = False
End Sub

Private Sub cmdExit_Click()
    UnloadForm Me
End Sub

Private Sub cmdFind_Click()
    MsgBox "Use the List view to find...", vbInformation, "Find"
End Sub

Private Sub cmdNext_Click()
    RSLOAN.MoveNext
    If RSLOAN.EOF Then
        RSLOAN.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    RSLOAN.MovePrevious
    If RSLOAN.BOF Then
        RSLOAN.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "FILES LOAD CODE") = False Then Exit Sub

    Screen.MousePointer = 11
    rptLoan.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
    rptLoan.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
    rptLoan.Formulas(2) = "COMPANYTIN = '" & COMPANY_TIN & "'"
    rptLoan.Formulas(3) = "PRINTBY = '" & LOGNAME & "'"

    PrintSQLReport rptLoan, HRMS_REPORT_PATH & "Loan Code.rpt", "", DMIS_REPORT_Connection, 1

    LogAudit "V", "PRINT LOAN CODE", ""
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    'On Error GoTo Errorcode
    Dim LONCODE                                                       As String
    Dim LONDESC                                                       As String


    If txtCode.Text = "" Then MsgBox "Loan Code cannot be Blank", vbInformation, "Incomplete Data": txtCode.SetFocus: Exit Sub
    If txtDesc.Text = "" Then MsgBox "Loan Description Cannot be Blank", vbInformation, "Incomplete Data": txtDesc.SetFocus: Exit Sub

    LONCODE = N2Str2Null(txtCode.Text)
    LONDESC = N2Str2Null(txtDesc.Text)
    If ADDOREDIT = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_LoanCode (code,Description) " & _
                       " values (" & LONCODE & ", " & LONDESC & ")"

        LogAudit "A", "ADD LOAN CODE", LONCODE
        ShowSuccessFullyAdded
    Else
'        gconDMIS.Execute "update HRMS_LoanCode set" & _
'                       " description = " & LONDESC & _
'                       " description = " & LONDESC & _
'                       " where id = " & labid.Caption

        gconDMIS.Execute "update HRMS_LoanCode set" & _
                       " description = " & LONDESC & _
                       " where id = " & labid.Caption


        LogAudit "E", "UPDATE LOAN CODE", LONCODE
        ShowSuccessFullyUpdated
    End If

    rsrefresh
    DisplayAllLoanCode
    cmdCancel.Value = True
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode

    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    DisplayAllLoanCode
    rsrefresh
    StoreMemVars

End Sub

Private Sub lstLoan_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    labid = lstLoan.SelectedItem.ListSubItems(2).Text
    RSLOAN.MoveFirst
    RSLOAN.Find ("ID=" & lstLoan.SelectedItem.ListSubItems(2).Text)
    StoreMemVars


End Sub


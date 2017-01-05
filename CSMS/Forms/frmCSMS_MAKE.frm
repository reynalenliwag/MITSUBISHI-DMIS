VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMS_MAKE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Make MasterFile"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_MAKE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   5265
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   30
      ScaleHeight     =   3255
      ScaleWidth      =   5325
      TabIndex        =   10
      Top             =   1620
      Width           =   5325
      Begin MSComctlLib.ListView lsvMAKE 
         Height          =   3045
         Left            =   60
         TabIndex        =   11
         Top             =   60
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   5371
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "MAKE"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "FLAT RATE"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   90
      ScaleHeight     =   1695
      ScaleWidth      =   4845
      TabIndex        =   7
      Top             =   30
      Width           =   4845
      Begin Crystal.CrystalReport rptMAKE 
         Left            =   4320
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox TXTFLATRATE 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   990
         TabIndex        =   2
         Top             =   1020
         Width           =   1665
      End
      Begin VB.TextBox TXTMAKE 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   990
         TabIndex        =   1
         Top             =   570
         Width           =   3405
      End
      Begin VB.TextBox TXTCODE 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   990
         MaxLength       =   1
         TabIndex        =   0
         Top             =   120
         Width           =   795
      End
      Begin VB.Label labid 
         Height          =   345
         Left            =   4200
         TabIndex        =   21
         Top             =   90
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label LBLCAP 
         AutoSize        =   -1  'True
         Caption         =   "Flatrate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   1140
         Width           =   540
      End
      Begin VB.Label LBLCAP 
         AutoSize        =   -1  'True
         Caption         =   "Make"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   435
         TabIndex        =   19
         Top             =   690
         Width           =   375
      End
      Begin VB.Label LBLCAP 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   450
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   -6720
      ScaleHeight     =   960
      ScaleWidth      =   12015
      TabIndex        =   12
      Top             =   4800
      Width           =   12015
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
         Left            =   11220
         MouseIcon       =   "frmCSMS_MAKE.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_MAKE.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exit Window"
         Top             =   75
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
         Left            =   10530
         MouseIcon       =   "frmCSMS_MAKE.frx":153A
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_MAKE.frx":168C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Print this Record"
         Top             =   75
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
         Left            =   9840
         MouseIcon       =   "frmCSMS_MAKE.frx":19F2
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_MAKE.frx":1B44
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Delete Selected Record"
         Top             =   75
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
         Left            =   9150
         MouseIcon       =   "frmCSMS_MAKE.frx":1E6F
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_MAKE.frx":1FC1
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Edit Selected Record"
         Top             =   75
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
         Left            =   8460
         MouseIcon       =   "frmCSMS_MAKE.frx":231D
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_MAKE.frx":246F
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Add Record"
         Top             =   75
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
         Left            =   7770
         MouseIcon       =   "frmCSMS_MAKE.frx":2782
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_MAKE.frx":28D4
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Move to Next Record"
         Top             =   75
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
         Left            =   7080
         MouseIcon       =   "frmCSMS_MAKE.frx":2C2C
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_MAKE.frx":2D7E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Move to Previous Record"
         Top             =   75
         Width           =   705
      End
      Begin VB.Label lblSTATUS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   3435
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   3810
      ScaleHeight     =   885
      ScaleWidth      =   1590
      TabIndex        =   15
      Top             =   4830
      Width           =   1590
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
         MouseIcon       =   "frmCSMS_MAKE.frx":30DD
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_MAKE.frx":322F
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Cancel"
         Top             =   75
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
         MouseIcon       =   "frmCSMS_MAKE.frx":356D
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_MAKE.frx":36BF
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Save this Record"
         Top             =   75
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmCSMS_MAKE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMAKE                                             As ADODB.Recordset
Dim ADD_OR_EDIT                                        As String

Sub EnabledPic(COND As Boolean)
    picAdds.Visible = COND
    picSaves.Visible = Not COND
    Picture1.Enabled = Not COND
    Picture2.Enabled = COND
End Sub

Sub initMemvars()
    txtCode.Text = ""
    txtMake.Text = ""
    txtflatrate.Text = ""
End Sub

Sub rsRefresh()
    Set rsMAKE = New ADODB.Recordset
    rsMAKE.Open "SELECT * FROM ALL_MAKE ORDER BY MAKE", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not (rsMAKE.BOF And rsMAKE.EOF) Then
        txtCode.Text = Null2String(rsMAKE!Code)
        txtMake.Text = Null2String(rsMAKE!Make)
        txtflatrate.Text = Null2String(rsMAKE!FLATRATE)
        labID.Caption = rsMAKE!ID
    Else
        ShowNoRecord
        cmdAdd_Click
    End If
End Sub

Sub FillGrid()
    Dim rstmp                                          As New ADODB.Recordset
    Dim Item                                           As ListItem

    Set rstmp = gconDMIS.Execute("SELECT * FROM ALL_MAKE ORDER BY MAKE")
    lsvMAKE.ListItems.Clear
    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            Set Item = lsvMAKE.ListItems.Add(, , Null2String(rstmp!Code))
            Item.SubItems(1) = Null2String(rstmp!Make)
            Item.SubItems(2) = Null2String(rstmp!FLATRATE)
            Item.SubItems(3) = rstmp!ID

            rstmp.MoveNext
        Loop
    End If

    Set rstmp = Nothing
End Sub

Private Sub cmdAdd_Click()
    On Error Resume Next
    If Function_Access(LOGID, "Acess_ADD", "MAKE") = False Then Exit Sub
    ADD_OR_EDIT = "ADD"

    initMemvars
    EnabledPic False
    txtCode.SetFocus
End Sub

Private Sub cmdCancel_Click()
    EnabledPic True
    rsMAKE.Find "ID=" & labID.Caption & ""
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "MAKE") = False Then Exit Sub

    If lsvMAKE.ListItems.Count = 0 Then
        ShowNothingToDeleteMsg
        Exit Sub
    End If
    If MsgBox("Delete This Record", vbQuestion + vbYesNo, "Are You Sure") = False Then Exit Sub

    SQL_STATEMENT = "DELETE FROM ALL_MAKE WHERE ID = " & labID.Caption & ""
    gconDMIS.Execute SQL_STATEMENT

    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("X", "MAKE", SQL_STATEMENT, labID, "", "CODE: " & txtCode, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    ShowDeletedMsg
    FillGrid
    rsRefresh
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "MAKE") = False Then Exit Sub
    ADD_OR_EDIT = "EDIT"

    EnabledPic False
    txtCode.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rsMAKE.MoveNext
    If rsMAKE.EOF Then
        rsMAKE.MoveLast
        Call ShowLastRecordMsg
    End If
    Call StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsMAKE.MovePrevious
    If rsMAKE.BOF Then
        rsMAKE.MoveFirst
        Call ShowLastRecordMsg
    End If
    Call StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_print", "MAKE") = False Then Exit Sub

    Screen.MousePointer = 11
    rptMAKE.ReportTitle = "Vehicle Delear Make Master file"
    rptMAKE.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptMAKE.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptMAKE.Formulas(2) = "PrintedBy = '" & LOGNAME & "'"

    PrintSQLReport rptMAKE, CSMS_REPORT_PATH & "Make.rpt", "", CSMS_REPORT_CONNECTION, 1
    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("V", "MAKE", "", labID, "", "CODE: " & txtCode, "", "")
    'NEW LOG AUDIT-----------------------------------------------------
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    If txtCode.Text = "" Then
        ShowIsRequiredMsg "Code Cannot be Blank"
        txtCode.SetFocus
        Exit Sub
    End If

    If txtMake.Text = "" Then
        ShowIsRequiredMsg "Make Cannot be Blank"
        txtMake.SetFocus
        Exit Sub
    End If

    If txtflatrate.Text = "" Then
        ShowIsRequiredMsg "Flat Rate Cannot be Blank"
        txtflatrate.SetFocus
        Exit Sub
    End If

    Dim rstmp                                          As New ADODB.Recordset
    If ADD_OR_EDIT = "ADD" Then
        Set rstmp = gconDMIS.Execute("SELECT CODE FROM ALL_MAKE WHERE CODE = '" & txtCode.Text & "'")
        If Not (rstmp.BOF And rstmp.EOF) Then
            MsgBox "Code Already Exist", vbInformation, "CSMS"
            txtCode.SetFocus
            Exit Sub
        End If
    Else
        Set rstmp = gconDMIS.Execute("SELECT id,CODE FROM ALL_MAKE WHERE CODE = '" & txtCode.Text & "'")
        If Not (rstmp.BOF And rstmp.EOF) Then
            If Not labID.Caption = rstmp!ID Then
                MsgBox "Code Already Exist", vbInformation, "CSMS"
                txtCode.SetFocus
                Exit Sub
            End If
        End If
    End If

    Dim VMAKE                                          As String
    Dim VCODE                                          As String
    Dim VFLATRATE                                      As String

    VCODE = N2Str2Null(txtCode)
    VMAKE = N2Str2Null(txtMake)
    VFLATRATE = N2Str2Null(txtflatrate)

    If ADD_OR_EDIT = "ADD" Then
        SQL_STATEMENT = "INSERT INTO ALL_MAKE (CODE,MAKE,FLATRATE) VALUES(" & VCODE & "," & VMAKE & "," & VFLATRATE & ")"
        gconDMIS.Execute SQL_STATEMENT
        Set rstmp = gconDMIS.Execute("SELECT ID FROM ALL_MAKE WHERE CODE = '" & txtCode & "'")
        If Not (rstmp.BOF And rstmp.EOF) Then
            labID = rstmp!ID
        End If

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("A", "MAKE", SQL_STATEMENT, labID, "", "CODE: " & txtCode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "UPDATE ALL_MAKE SET CODE = " & VCODE & _
                        ",MAKE = " & VMAKE & _
                        ",FLATRATE = " & VFLATRATE & " WHERE ID = " & labID.Caption & ""
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "MAKE", SQL_STATEMENT, labID, "", "CODE: " & txtCode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        ShowSuccessFullyUpdated
    End If

    FillGrid
    rsRefresh
    rsMAKE.Find "ID=" & labID.Caption & ""
    cmdCancel_Click
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (MAKE MASTER FILE)"
            Call frmALL_AuditInquiry.DisplayHistory(labID, "MAKE", "")

    End Select
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)

    initMemvars
    FillGrid
    rsRefresh
    StoreMemVars
End Sub

Private Sub lsvMAKE_DblClick()
    If lsvMAKE.ListItems.Count = 0 Then Exit Sub

    cmdEdit_Click
End Sub

Private Sub lsvMAKE_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtCode.Text = Item.Text
    txtMake.Text = Item.ListSubItems(1)
    txtflatrate.Text = Item.ListSubItems(2)
    labID.Caption = Item.ListSubItems(3)
End Sub


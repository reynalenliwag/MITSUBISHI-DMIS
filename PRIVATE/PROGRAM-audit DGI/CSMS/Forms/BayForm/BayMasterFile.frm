VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBayMasterFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bay Location Data Entry"
   ClientHeight    =   5370
   ClientLeft      =   75
   ClientTop       =   435
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "BayMasterFile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   5835
   Begin VB.PictureBox picAdds 
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
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   6075
      TabIndex        =   12
      Top             =   4410
      Width           =   6075
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   5010
         MouseIcon       =   "BayMasterFile.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "BayMasterFile.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Exit Window"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   4320
         MouseIcon       =   "BayMasterFile.frx":153A
         MousePointer    =   99  'Custom
         Picture         =   "BayMasterFile.frx":168C
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Delete Selected Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   3630
         MouseIcon       =   "BayMasterFile.frx":19B7
         MousePointer    =   99  'Custom
         Picture         =   "BayMasterFile.frx":1B09
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   2940
         MouseIcon       =   "BayMasterFile.frx":1E65
         MousePointer    =   99  'Custom
         Picture         =   "BayMasterFile.frx":1FB7
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Add Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   2250
         MouseIcon       =   "BayMasterFile.frx":22CA
         MousePointer    =   99  'Custom
         Picture         =   "BayMasterFile.frx":241C
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Find a Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   1560
         MouseIcon       =   "BayMasterFile.frx":2716
         MousePointer    =   99  'Custom
         Picture         =   "BayMasterFile.frx":2868
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Move to Next Record"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   870
         MouseIcon       =   "BayMasterFile.frx":2BC0
         MousePointer    =   99  'Custom
         Picture         =   "BayMasterFile.frx":2D12
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   705
      End
   End
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
      Height          =   1095
      Left            =   30
      TabIndex        =   1
      Top             =   -60
      Width           =   5715
      Begin VB.TextBox txtdesc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   375
         Left            =   1200
         TabIndex        =   23
         Top             =   630
         Width           =   4425
      End
      Begin VB.TextBox txtcode 
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
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
         Height          =   345
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   0
         Top             =   180
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   225
         Left            =   675
         TabIndex        =   3
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Height          =   225
         Left            =   180
         TabIndex        =   2
         Top             =   690
         Width           =   945
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
      Height          =   3375
      Left            =   30
      TabIndex        =   6
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   7
         Top             =   570
         Width           =   5535
      End
      Begin MSComctlLib.ListView lstbay 
         Height          =   2325
         Left            =   60
         TabIndex        =   8
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
         MouseIcon       =   "BayMasterFile.frx":3071
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
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
         TabIndex        =   9
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.PictureBox picSaves 
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
      Left            =   4260
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   20
      Top             =   4455
      Width           =   1800
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   750
         MouseIcon       =   "BayMasterFile.frx":31D3
         MousePointer    =   99  'Custom
         Picture         =   "BayMasterFile.frx":3325
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   60
         MouseIcon       =   "BayMasterFile.frx":3663
         MousePointer    =   99  'Custom
         Picture         =   "BayMasterFile.frx":37B5
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label labPrev 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8160
      TabIndex        =   5
      Top             =   570
      Width           =   195
   End
   Begin VB.Label labid 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   4
      Top             =   690
      Width           =   225
   End
End
Attribute VB_Name = "frmBayMasterFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UPDATE_MODE                                        As Boolean
Dim RS                                                 As New ADODB.Recordset

Sub initMemvars()
    txtcode.Text = ""
    txtdesc.Text = ""
End Sub

Sub Saveinformation()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim thecode                                        As String
    Dim theDescription                                 As String
    Dim thebayStatus                                   As String
    Dim theRo                                          As String
    Dim theRostatus                                    As String
    Dim thePlateNo                                     As String


    
    
    ' JBF 05/31/2010
    ' Validation to avoid saving blank
    ' *******************************************
    
    If txtcode.Text = "" Then
       MsgBox "Code cannot be blank.", vbExclamation, "WARNING"
       On Error Resume Next
       txtcode.SetFocus
       Exit Sub
    End If
    
    If txtdesc.Text = "" Then
        MsgBox "Description cannot be blank.", vbInformation, "WARNING"
        On Error Resume Next
        txtdesc.SetFocus
        Exit Sub
     End If
       
    ' *******************************************
    
    
    thecode = N2Str2Null(txtcode.Text)
    theDescription = N2Str2Null(txtdesc.Text)
    thebayStatus = N2Str2Null("Available")
    theRo = N2Str2Null("")
    theRostatus = N2Str2Null("")
    thePlateNo = N2Str2Null("")
    
    
    
    
    
If NumericVal(txtcode.Text) > 30 Then
    MsgBox "Code is greater than the default Bay..", vbInformation, "Info"
    Exit Sub
End If
    Set RS = gconDMIS.Execute("Select count(*)bilang from CSMS_baymonitoring where bay_code='" & (txtcode.Text) & "'")
    If RS!BILANG = 1 Then
            MsgBox "Code is already exist..", vbExclamation, "WARNING"
    Exit Sub
    End If
    Set RS = Nothing
    If UPDATE_MODE = False Then
        SQL = "Insert into CSMS_Baymonitoring(Bay_Code,bay_description,bay_status,RO,RO_status,Plate_no) VALUES(" & thecode & _
              "," & theDescription & "," & thebayStatus & "," & theRo & "," & theRostatus & "," & thePlateNo & ")"
        gconDMIS.Execute SQL

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("A", "BAY MASTER FILE", SQL, labid, "", "CODE : " & txtcode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        ShowSuccessFullyAdded
    Else
        SQL = "UPDATE CSMS_Baymonitoring set bay_description=" & theDescription & " where bay_code=" & thecode & ""
        gconDMIS.Execute SQL

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("E", "BAY MASTER FILE", SQL, labid, "", "CODE : " & txtcode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        ShowSuccessFullyUpdated
    End If


    cmdCancel.Value = True
    initMemvars
    DisplayInformation
    StoreMemVars
End Sub

Sub DisplayInformation()
    Dim Item                                           As ListItem
    Dim RS                                             As New ADODB.Recordset
    Dim SQL                                            As String

    SQL = "SELECT bay_code,Bay_description from CSMS_BayMonitoring order by bay_code asc"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    lstbay.ListItems.Clear

    Do While Not RS.EOF
        Set Item = lstbay.ListItems.Add(, , Null2String(RS!bay_code))
        Item.SubItems(1) = Null2String(RS!bay_description)
        RS.MoveNext
    Loop

    Set RS = Nothing
End Sub

Sub rsRefresh()
    Set RS = New ADODB.Recordset
    Call RS.Open("SELECT * FROM CSMS_BayMonitoring", gconDMIS, adOpenKeyset, adLockReadOnly)
End Sub

Sub StoreMemVars()
    If Not RS.EOF And Not RS.BOF Then
        labid.Caption = (RS!ID)
        txtcode = Null2String(RS!bay_code)
        txtdesc = Null2String(RS!bay_description)
    End If
End Sub

Sub SearhME()
    Dim RS                                             As New ADODB.Recordset
    Dim Item                                           As ListItem
    Dim SQL                                            As String
    Dim Keyword                                        As String

    SQL = "SELECT * FROM CSMS_BAymonitoring WHERE"
    Keyword = Trim(txtSearch.Text)

    If Len(Keyword) = 0 Then
        DisplayInformation
        Exit Sub
    End If
    If optDesc.Value = True Then
        SQL = SQL & " Bay_description like '" & Keyword & "%'"
    ElseIf optCode.Value = True Then
        SQL = SQL & " Bay_code like '" & Keyword & "%'"
    End If

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    lstbay.ListItems.Clear

    Do While Not RS.EOF
        Set Item = lstbay.ListItems.Add(, , Null2String(RS!bay_code))
        Item.SubItems(1) = Null2String(RS!bay_description)
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "ACESS_ADD", "BAY MASTER FILE") = False Then Exit Sub
    picAdds.Visible = False
    Frame1.Enabled = True
    initMemvars
    UPDATE_MODE = False
End Sub

Private Sub cmdCancel_Click()
    picAdds.Visible = True
    Frame1.Enabled = False
    initMemvars
    UPDATE_MODE = False
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "ACESS_DELETE", "BAY MASTER FILE") = False Then Exit Sub

    If MsgBox("Are you sure you want to delete this Information", vbQuestion + vbYesNo) = vbYes Then
        SQL_STATEMENT = "DELETE From CSMS_Baymonitoring where bay_code = '" & txtcode & "'"
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("X", "", SQL_STATEMENT, labid, "", "CODE: " & txtcode, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        ShowDeletedMsg
        initMemvars
        DisplayInformation
    End If
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "ACESS_EDIT", "BAY MASTER FILE") = False Then Exit Sub
    UPDATE_MODE = True
    Frame1.Enabled = True
    picAdds.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    RS.MoveNext
    If RS.EOF Then
        RS.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    RS.MovePrevious
    If RS.BOF Then
        RS.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    Saveinformation
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            If picAdds.Visible = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (BAY MASTER FILE)"
            Call frmALL_AuditInquiry.DisplayHistory(labid, "BAY MASTER FILE", "")

    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Frame1.Enabled = False
    UPDATE_MODE = False
    DisplayInformation
    rsRefresh
    StoreMemVars
End Sub

Private Sub lstbay_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtcode = lstbay.ListItems(lstbay.SelectedItem.Index)
    txtdesc = lstbay.SelectedItem.SubItems(1)
End Sub

Private Sub txtcode_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
    Else
    End If
End Sub

Private Sub txtcode_Validate(Cancel As Boolean)
    If IsNumeric(txtcode.Text) = False Then
        MsgBox "Please input integer Value from 1 to 30 only..", vbInformation, "Info"
        Cancel = True
    End If
End Sub

Private Sub txtSearch_Change()
    SearhME
End Sub

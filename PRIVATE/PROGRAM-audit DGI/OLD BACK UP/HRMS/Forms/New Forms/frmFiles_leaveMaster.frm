VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmHRMSFiles_leaveMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leave Master File"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFiles_leaveMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   6240
   Begin VB.PictureBox picAdd 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   660
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   12
      Top             =   5190
      Width           =   5580
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   4860
         MouseIcon       =   "frmFiles_leaveMaster.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "frmFiles_leaveMaster.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   4170
         MouseIcon       =   "frmFiles_leaveMaster.frx":0A42
         MousePointer    =   99  'Custom
         Picture         =   "frmFiles_leaveMaster.frx":0B94
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   3480
         MouseIcon       =   "frmFiles_leaveMaster.frx":0EBF
         MousePointer    =   99  'Custom
         Picture         =   "frmFiles_leaveMaster.frx":1011
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   2790
         MouseIcon       =   "frmFiles_leaveMaster.frx":136D
         MousePointer    =   99  'Custom
         Picture         =   "frmFiles_leaveMaster.frx":14BF
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton k 
         Caption         =   "&Next"
         Height          =   795
         Left            =   2100
         MouseIcon       =   "frmFiles_leaveMaster.frx":17D2
         MousePointer    =   99  'Custom
         Picture         =   "frmFiles_leaveMaster.frx":1924
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Move to Next Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   1410
         MouseIcon       =   "frmFiles_leaveMaster.frx":1C7C
         MousePointer    =   99  'Custom
         Picture         =   "frmFiles_leaveMaster.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1965
      Left            =   30
      ScaleHeight     =   1965
      ScaleWidth      =   6165
      TabIndex        =   11
      Top             =   90
      Width           =   6165
      Begin VB.OptionButton Option1 
         Caption         =   "# of leaves accumulated every month"
         Height          =   255
         Left            =   1530
         TabIndex        =   21
         Top             =   1410
         Width           =   3045
      End
      Begin VB.OptionButton Option2 
         Caption         =   "# of leaves fixed in a year"
         Height          =   345
         Left            =   1530
         TabIndex        =   20
         Top             =   1650
         Width           =   2985
      End
      Begin VB.TextBox txtLeaveAvail 
         Height          =   345
         Left            =   1530
         TabIndex        =   2
         Top             =   1020
         Width           =   1005
      End
      Begin VB.TextBox txtLeaveDesc 
         Height          =   345
         Left            =   1530
         TabIndex        =   1
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox txtLeaveCode 
         Height          =   345
         Left            =   1530
         TabIndex        =   0
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label labId 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   4440
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Available Leave"
         Height          =   210
         Index           =   2
         Left            =   300
         TabIndex        =   18
         Top             =   1110
         Width           =   1155
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Leave Description"
         Height          =   210
         Index           =   1
         Left            =   150
         TabIndex        =   17
         Top             =   690
         Width           =   1305
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Leave Code"
         Height          =   210
         Index           =   0
         Left            =   585
         TabIndex        =   16
         Top             =   300
         Width           =   870
      End
   End
   Begin VB.PictureBox picSearch 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   30
      ScaleHeight     =   3015
      ScaleWidth      =   6195
      TabIndex        =   10
      Top             =   2100
      Width           =   6195
      Begin MSComctlLib.ListView lsvDet 
         Height          =   2985
         Left            =   60
         TabIndex        =   3
         Top             =   30
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   5265
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
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "No of Leave Avial."
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Type"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "id"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.PictureBox picSave 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4800
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   13
      Top             =   5190
      Width           =   1440
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   720
         MouseIcon       =   "frmFiles_leaveMaster.frx":212D
         MousePointer    =   99  'Custom
         Picture         =   "frmFiles_leaveMaster.frx":227F
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   30
         MouseIcon       =   "frmFiles_leaveMaster.frx":25BD
         MousePointer    =   99  'Custom
         Picture         =   "frmFiles_leaveMaster.frx":270F
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmHRMSFiles_leaveMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLeave As ADODB.Recordset
Dim ADDOREDIT As String

Private Sub cmdAdd_Click()
    ADDOREDIT = "ADD"
    
    picAdd.Visible = False
    picSave.Visible = True
    picSearch.Enabled = False
    picMain.Enabled = True
    
    Call InitMemVars
    On Error Resume Next
    txtLeaveCode.SetFocus
End Sub

Private Sub cmdCancel_Click()
    picMain.Enabled = False
    picSearch.Enabled = True
    picSave.Visible = False
    picAdd.Visible = True
    
    Call StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Delete this Leave Type, are you sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    
    
    gconDMIS.Execute ("DELETE FROM HRMS_LeaveMaster WHERE ID = " & labId & "")
    Call ShowDeletedMsg
    
    Call rsRefresh
    Call DisplayGrid
    Call StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    ADDOREDIT = "EDIT"
    
    picMain.Enabled = True
    picSearch.Enabled = False
    picAdd.Visible = False
    picSave.Visible = True
    
    On Error Resume Next
    txtLeaveCode.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim xLEAVECODE As String
    Dim xLEAVEDESC As String
    Dim xCOUNT As Integer
    Dim xLeave_Type As Integer
    Dim RSTMP As New ADODB.Recordset
    
    If Option1.Value = True Then
        xLeave_Type = 1
    Else
        xLeave_Type = 0
    End If
    
    xLEAVECODE = N2Str2Null(txtLeaveCode)
    xLEAVEDESC = N2Str2Null(txtLeaveDesc)
    xCOUNT = NumericVal(txtLeaveAvail)
    
    If ADDOREDIT = "ADD" Then
        Set RSTMP = gconDMIS.Execute("SELECT LEAVE_CODE FROM HRMS_LeaveMaster WHERE LEAVE_CODE = " & xLEAVECODE & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            MsgBox "Leave Code already Exist in the Master File", vbExclamation, "Duplicate Code"
            On Error Resume Next
            txtLeaveCode.SetFocus
            Exit Sub
        End If
        Set RSTMP = gconDMIS.Execute("SELECT LEAVE_DESC FROM HRMS_LeaveMaster WHERE LEAVE_CODE = " & xLEAVECODE & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            MsgBox "Leave Description already Exist in the Master File", vbExclamation, "Duplicate Description"
            On Error Resume Next
            txtLeaveDesc.SetFocus
            Exit Sub
        End If
    Else
        Set RSTMP = gconDMIS.Execute("SELECT ID, LEAVE_CODE FROM HRMS_LeaveMaster WHERE LEAVE_CODE = " & xLEAVECODE & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            If Not labId = RSTMP!ID Then
                MsgBox "Leave Code already Exist in the Master File", vbExclamation, "Duplicate Code"
                On Error Resume Next
                txtLeaveCode.SetFocus
                Exit Sub
            End If
        End If
        Set RSTMP = gconDMIS.Execute("SELECT ID, LEAVE_DESC FROM HRMS_LeaveMaster WHERE LEAVE_DESC = " & xLEAVEDESC & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            If Not labId = RSTMP!ID Then
                MsgBox "Leave Description already Exist in the Master File", vbExclamation, "Duplicate Description"
                On Error Resume Next
                txtLeaveDesc.SetFocus
                Exit Sub
            End If
        End If
    End If
    Set RSTMP = Nothing
    
    If ADDOREDIT = "ADD" Then
        gconDMIS.Execute ("INSERT INTO HRMS_LeaveMaster (LEAVE_CODE, LEAVE_DESC, ISSTANDARD, TYPE, DAYS_NO) " & _
            " VALUES(" & xLEAVECODE & _
            ", " & xLEAVEDESC & ", " & xLeave_Type & _
            ", 'TRUE' " & _
            ", " & xCOUNT & ")")
        
        Set RSTMP = New ADODB.Recordset
        Set RSTMP = gconDMIS.Execute("SELECT ID FROM HRMS_LeaveMaster WHERE LEAVE_CODE = " & xLEAVECODE & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            labId = RSTMP!ID
        End If
        Set RSTMP = Nothing
        
        ShowSuccessFullyAdded
    Else
        gconDMIS.Execute ("UPDATE HRMS_LeaveMaster SET " & _
            " LEAVE_CODE = " & xLEAVECODE & _
            ", LEAVE_DESC = " & xLEAVEDESC & _
            ", ISSTANDARD = 'TRUE', TYPE = " & xLeave_Type & _
            ", DAYS_NO = " & xCOUNT & _
            " WHERE ID = " & labId & "")
            
        ShowSuccessFullyUpdated
    End If
    
    Call rsRefresh
    Call DisplayGrid
    rsLeave.Find "ID = " & labId
    
    Call cmdCancel_Click
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    
    Call rsRefresh
    Call DisplayGrid
    Call StoreMemVars
End Sub

Sub InitMemVars()
    txtLeaveCode = ""
    txtLeaveDesc = ""
    txtLeaveAvail = ""
End Sub

Sub rsRefresh()
    Set rsLeave = New ADODB.Recordset
    rsLeave.Open "SELECT * FROM HRMS_LeaveMaster ORDER BY LEAVE_CODE", gconDMIS
End Sub

Sub StoreMemVars()
    If Not (rsLeave.BOF And rsLeave.EOF) Then
        labId.Caption = rsLeave!ID
        txtLeaveCode = Null2String(rsLeave!LEAVE_CODE)
        txtLeaveDesc = Null2String(rsLeave!LEAVE_DESC)
        txtLeaveAvail = NumericVal(rsLeave!DAYS_NO)
        
        If Null2Bool(rsLeave!Type) = False Then
            Option2.Value = True
            Option1.Value = False
        Else
            Option1.Value = True
            Option2.Value = False
        End If
    Else
        ShowNoRecord
        Call cmdAdd_Click
    End If
End Sub

Sub DisplayGrid()
    Dim Item As ListItem
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_LeaveMaster ORDER BY LEAVE_cODE")
    lsvDet.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set Item = lsvDet.ListItems.Add(, , Null2String(RSTMP!LEAVE_CODE))
            Item.SubItems(1) = Null2String(RSTMP!LEAVE_DESC)
            Item.SubItems(2) = NumericVal(RSTMP!DAYS_NO)
            If Null2Bool(RSTMP!Type) = True Then
                Item.SubItems(3) = 1
            Else
                Item.SubItems(3) = 0
            End If
            Item.SubItems(4) = RSTMP!ID
            
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Public Function LimitChar(ByVal alpha As String, ByVal k As Integer)
    If InStr(alpha, Chr(k)) > 0 Or k = 8 Then
        LimitChar = k
    Else
        LimitChar = 0
    End If
End Function

Private Sub lsvDet_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_LeaveMaster WHERE ID = " & Item.SubItems(4) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        labId.Caption = RSTMP!ID
        txtLeaveCode = Null2String(RSTMP!LEAVE_CODE)
        txtLeaveDesc = Null2String(RSTMP!LEAVE_DESC)
        txtLeaveAvail = NumericVal(RSTMP!DAYS_NO)
        
        If Null2Bool(RSTMP!Type) = False Then
            Option2.Value = True
            Option1.Value = False
        Else
            Option1.Value = True
            Option2.Value = False
        End If
    End If
    Set RSTMP = Nothing
End Sub

Private Sub txtLeaveAvail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

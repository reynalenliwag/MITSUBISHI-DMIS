VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMSSelectContractor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Contractor"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   BeginProperty Font 
      Name            =   "Tunga"
      Size            =   9.75
      Charset         =   1
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelectContractor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   5070
      MouseIcon       =   "frmSelectContractor.frx":06D2
      MousePointer    =   99  'Custom
      Picture         =   "frmSelectContractor.frx":0824
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Close Window"
      Top             =   2880
      Width           =   735
   End
   Begin MSComctlLib.ListView listContractor 
      Height          =   2775
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "First Name"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "LastName"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "mname"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Company Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Address"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.PictureBox picselect 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   60
      ScaleHeight     =   2745
      ScaleWidth      =   5715
      TabIndex        =   2
      Top             =   60
      Width           =   5745
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   -30
         ScaleHeight     =   105
         ScaleWidth      =   5715
         TabIndex        =   3
         Top             =   2640
         Width           =   5745
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   285
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   5745
         _Version        =   655364
         _ExtentX        =   10134
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "CONTRACTOR INFORMATION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16711680
         GradientColorDark=   16711680
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Esc - Cancel"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   1005
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Press Enter key to Assign Contractor"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   2565
         TabIndex        =   14
         Top             =   2250
         Width           =   3060
      End
      Begin VB.Label lbladdress 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1515
         TabIndex        =   13
         Top             =   1170
         Width           =   4095
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1515
         TabIndex        =   12
         Top             =   780
         Width           =   4095
      End
      Begin VB.Label lbllastname 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6645
         TabIndex        =   11
         Top             =   840
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Label lblfirstname 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6645
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Label lblcode 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1515
         TabIndex        =   9
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Address"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   1260
         Width           =   1305
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   300
         TabIndex        =   7
         Top             =   900
         Width           =   1125
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LastName"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5820
         TabIndex        =   6
         Top             =   930
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FirstName"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5790
         TabIndex        =   5
         Top             =   540
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1050
         TabIndex        =   4
         Top             =   450
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Assigned"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4350
      MouseIcon       =   "frmSelectContractor.frx":0B62
      MousePointer    =   99  'Custom
      Picture         =   "frmSelectContractor.frx":0CB4
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Assign Technician"
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblCONCODE 
      AutoSize        =   -1  'True
      Caption         =   "con code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2430
      TabIndex        =   25
      Top             =   5220
      Width           =   840
   End
   Begin VB.Label lblLineNo 
      Caption         =   "LINE NO"
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
      Left            =   3150
      TabIndex        =   23
      Top             =   4380
      Width           =   1815
   End
   Begin VB.Label labCust 
      Caption         =   "labCust"
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
      Left            =   330
      TabIndex        =   22
      Top             =   5850
      Width           =   3135
   End
   Begin VB.Label labItemNo 
      Caption         =   "Item No"
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
      Left            =   330
      TabIndex        =   21
      Top             =   5460
      Width           =   3135
   End
   Begin VB.Label Label9 
      Caption         =   "Double Click to Select Contractor"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   20
      Top             =   2850
      Width           =   3255
   End
   Begin VB.Label lblplate 
      Caption         =   "theplate"
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
      Left            =   330
      TabIndex        =   19
      Top             =   5100
      Width           =   3135
   End
   Begin VB.Label lblModel 
      Caption         =   "model"
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
      Left            =   300
      TabIndex        =   18
      Top             =   4860
      Width           =   3135
   End
   Begin VB.Label lblCustomer 
      Caption         =   "thecus"
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
      Left            =   330
      TabIndex        =   17
      Top             =   4590
      Width           =   3135
   End
   Begin VB.Label lblRO 
      Caption         =   "thero"
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
      Left            =   450
      TabIndex        =   15
      Top             =   4260
      Width           =   1815
   End
End
Attribute VB_Name = "frmCSMSSelectContractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim thefirstname, thelastname, themiName, thecompany, theAddress As String
Attribute thelastname.VB_VarUserMemId = 1073938432
Attribute themiName.VB_VarUserMemId = 1073938432
Attribute thecompany.VB_VarUserMemId = 1073938432
Attribute theAddress.VB_VarUserMemId = 1073938432
Dim thecode                                            As String
Attribute thecode.VB_VarUserMemId = 1073938437
Dim rsRO_DET                                           As ADODB.Recordset
Attribute rsRO_DET.VB_VarUserMemId = 1073938438

Sub FillContractor()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim arnie                                          As ListItem

    SQL = "SELECT * FROM CSMS_Contractor"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    listContractor.ListItems.Clear

    Do While Not RS.EOF
        Set arnie = listContractor.ListItems.Add(, , Null2String(RS!Code))
        arnie.SubItems(1) = Null2String(RS!Firstname)
        arnie.SubItems(2) = Null2String(RS!lastname)
        arnie.SubItems(3) = Null2String(RS!MNAME)
        arnie.SubItems(4) = Null2String(RS!CompanyName)
        arnie.SubItems(5) = Null2String(RS!Address)
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Sub AssignContractor()
    If Not lblCONCODE.Caption = "" Then
        Dim RSTMP                                      As New ADODB.Recordset
        Dim RSTECH                                     As New ADODB.Recordset
        Dim VNULL                                      As String

        VNULL = N2Str2Null("")
        Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE LIVIL = '1' AND LINE_NO = '" & lblLINENO.Caption & "' AND REP_OR = '" & lblro.Caption & "'")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            If Null2String(RSTMP!DONE) = "W" Then
                MsgBox "You Cannot Assigned a Contractor, Job Is Already Started", vbExclamation
                Exit Sub
            ElseIf Null2String(RSTMP!DONE) = "Y" Then
                MsgBox "You Cannot Assigned a Contractor, Job Is Already Finished", vbExclamation
                Exit Sub
            Else
                Set RSTECH = gconDMIS.Execute("SELECT EMPNO FROM CSMS_VW_TECHNICIAN WHERE TECHNICIAN = '" & lblCONCODE.Caption & "'")
                If Not (RSTECH.BOF And RSTECH.EOF) Then
                    gconDMIS.Execute ("UPDATE HRMS_EMPINFO SET ASSIGNEDRO = " & VNULL & ",JSTATUS = 'A' where EMPNO = '" & RSTECH!EMPNO & "'")
                Else
                    gconDMIS.Execute ("UPDATE CSMS_CONTRACTORMONITORING SET ASSIGNEDRO = " & VNULL & ",STATUS = 'Available' where code = '" & lblCONCODE.Caption & "'")
                End If
            End If
        End If
    End If

    SQL_STATEMENT = "Update CSMS_Ro_det set " & _
                  " TECHNICIAN = '" & lblfirstname.Caption & "', TECHCODE = '" & lblcode & "', DONE = 'Y', STATUS = 'Y'" & _
                  " Where DETCDE = '" & LABITEMNO.Caption & "' AND REP_OR = '" & lblro.Caption & "' AND LINE_NO = '" & lblLINENO.Caption & "' And livil = '1'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("AS", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(lblro), "REP_OR", "CSMS_REPOR"), "", "JOB CODE: " & LABITEMNO & " - " & lblfirstname, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    'UpdateContractor

    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("Select STATUS from CSMS_RO_DET WHERE LIVIL = '1' AND REP_OR = '" & lblro.Caption & "' AND (DONE = 'N' OR DONE IS NULL OR DONE ='W')")

    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        'do nothing
    Else
        gconDMIS.Execute "update CSMS_RepairOrder set dateFinish = '" & LOGDATE & "', STATUS = 'Finish Job', JStatus = 'F' where RO_No = '" & lblro.Caption & "'"
    End If

    MessagePop InfoFriend, "RO Information Updated", "Contractor Succesfully Assigned to Job", 1000
    Call frmCSMS_ServiceCounter.Click_ScheduleGrid
    cmdClose.Value = True
End Sub

Sub UpdateContractor()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim thestatus                                      As String
    Dim vRO                                            As String

    vRO = lblro.Caption & "-" & lblLINENO.Caption
    thestatus = "Assigned"

    '    SQL = "Update CSMS_ContractorMonitoring set Status = '" & thestatus & _
         '        "',AssignedRO = '" & lblro.Caption & _
         '        "' Where Code = '" & lblcode.Caption & "'"

    SQL = "Update CSMS_ContractorMonitoring set Status = '" & thestatus & _
          "',AssignedRO = '" & vRO & _
          "' Where Code = '" & lblcode.Caption & "'"

    'Set rs = New ADODB.Recordset
    'Set rs = gconDMIS.Execute(SQL)

    gconDMIS.Execute (SQL)
End Sub

Private Sub cmdAdd_Click()
    thecode = listContractor.ListItems(listContractor.SelectedItem.Index).Text
    thefirstname = listContractor.SelectedItem.SubItems(1)
    thelastname = listContractor.SelectedItem.SubItems(2)
    themiName = listContractor.SelectedItem.SubItems(3)
    thecompany = listContractor.SelectedItem.SubItems(4)
    theAddress = listContractor.SelectedItem.SubItems(5)

    'LogAudit "A", "CONTRACTOR SELECTED", thelastname & " " & thefirstname

    lblcode.Caption = thecode
    lblfirstname.Caption = thefirstname
    lbllastname.Caption = thelastname
    lblCompany.Caption = thecompany
    lbladdress.Caption = theAddress
    If lblcode.Caption <> "" Then
        picselect.Visible = True
        picselect.ZOrder 0
    Else
        MsgBox "No Job Contractor Assigned!", vbCritical, "Empty"
    End If
    If MsgBox("Are you sure do you want to assinged contractor?", vbQuestion + vbYesNo) = vbYes Then
        AssignContractor
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ans                                            As String
    If picselect.Visible = True Then
        If KeyCode = vbKeyReturn Then
            ans = MsgBox("Are you sure do you wnt to assinged contractor?", vbQuestion + vbYesNo)
            If ans = vbYes Then
                AssignContractor
            End If
        End If
        If KeyCode = vbKeyEscape Then
            picselect.Visible = False
            picselect.ZOrder 1
        End If
    End If
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    FillContractor
    picselect.Visible = False
    picselect.ZOrder 1

    'If Not listContractor.ListItems.Count = 0 Then listContractor_DblClick
End Sub

Private Sub listContractor_DblClick()

    On Error Resume Next
    thecode = listContractor.ListItems(listContractor.SelectedItem.Index).Text
    thefirstname = listContractor.SelectedItem.SubItems(1)
    thelastname = listContractor.SelectedItem.SubItems(2)
    themiName = listContractor.SelectedItem.SubItems(3)
    thecompany = listContractor.SelectedItem.SubItems(4)
    theAddress = listContractor.SelectedItem.SubItems(5)
    LogAudit "A", "CONTRACTOR SELECTED", thelastname & " " & thefirstname

    lblcode.Caption = thecode
    lblfirstname.Caption = thefirstname
    lbllastname.Caption = thelastname
    lblCompany.Caption = thecompany
    lbladdress.Caption = theAddress
    If lblcode.Caption <> "" Then
        picselect.Visible = True
        picselect.ZOrder 0
    Else
        MsgBox "No Job Contractor Assigned!", vbCritical, "Empty"
    End If
End Sub


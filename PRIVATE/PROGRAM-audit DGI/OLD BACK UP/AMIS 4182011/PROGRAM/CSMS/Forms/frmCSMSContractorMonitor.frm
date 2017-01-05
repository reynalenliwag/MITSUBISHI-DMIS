VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMSContractorMonitor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contractor Monitoring"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11460
   Icon            =   "frmCSMSContractorMonitor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   11460
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   6330
      TabIndex        =   1
      Top             =   90
      Width           =   5115
      Begin VB.Frame TrapNoRO 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   1605
         Left            =   30
         TabIndex        =   12
         Top             =   390
         Width           =   4995
         Begin VB.PictureBox Picture7 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   2250
            Picture         =   "frmCSMSContractorMonitor.frx":0DCA
            ScaleHeight     =   735
            ScaleWidth      =   705
            TabIndex        =   13
            Top             =   390
            Width           =   705
         End
         Begin VB.Timer Timer3 
            Interval        =   500
            Left            =   90
            Top             =   540
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "No Repair Order "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   1140
            TabIndex        =   14
            Top             =   1140
            Width           =   2955
         End
      End
      Begin VB.TextBox txtplate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1050
         TabIndex        =   9
         Top             =   1590
         Width           =   1755
      End
      Begin VB.TextBox txtmodel 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         TabIndex        =   8
         Top             =   1170
         Width           =   1755
      End
      Begin VB.TextBox txtcust 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         TabIndex        =   7
         Top             =   780
         Width           =   3945
      End
      Begin VB.TextBox txtro 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1050
         TabIndex        =   6
         Top             =   420
         Width           =   1755
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   5115
         _Version        =   655364
         _ExtentX        =   9022
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "INFORMATION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16711680
         GradientColorDark=   16711680
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Plate No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   5
         Top             =   1650
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1230
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ro No"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView listContractor 
      Height          =   5355
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9446
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
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmCSMSContractorMonitor.frx":1676
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "code"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Contractor"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "R.Order"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   0
      EndProperty
   End
   Begin VB.PictureBox picin 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   6330
      ScaleHeight     =   3225
      ScaleWidth      =   5055
      TabIndex        =   10
      Top             =   2190
      Visible         =   0   'False
      Width           =   5085
      Begin VB.ComboBox cboOUT 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "frmCSMSContractorMonitor.frx":17D8
         Left            =   1050
         List            =   "frmCSMSContractorMonitor.frx":17E2
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1890
         Visible         =   0   'False
         Width           =   3105
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   5115
         Begin VB.TextBox txtContractorName 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1890
            TabIndex        =   20
            Top             =   450
            Width           =   3045
         End
         Begin VB.TextBox txtaddress 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1890
            TabIndex        =   19
            Top             =   810
            Width           =   3045
         End
         Begin VB.TextBox txtcompany 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1890
            TabIndex        =   18
            Top             =   1170
            Width           =   3045
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   285
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   5085
            _Version        =   655364
            _ExtentX        =   8969
            _ExtentY        =   503
            _StockProps     =   14
            Caption         =   "CONTRACTOR INFORMATION"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientColorLight=   16711680
            GradientColorDark=   16711680
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Full Name"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1020
            TabIndex        =   23
            Top             =   510
            Width           =   795
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Company Address"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   375
            TabIndex        =   22
            Top             =   900
            Width           =   1440
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contractor Company"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   105
            TabIndex        =   21
            Top             =   1260
            Width           =   1710
         End
      End
      Begin VB.CommandButton cmdOutCancel 
         Height          =   585
         Left            =   2580
         Picture         =   "frmCSMSContractorMonitor.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancel "
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdOut 
         Height          =   585
         Left            =   1020
         Picture         =   "frmCSMSContractorMonitor.frx":4D7F
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Clock Out Technician"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton cmdIn 
         Height          =   585
         Left            =   1020
         Picture         =   "frmCSMSContractorMonitor.frx":8815
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Clock In Technician"
         Top             =   2520
         Width           =   1575
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   27
      Top             =   90
      Width           =   6135
      _Version        =   655364
      _ExtentX        =   10821
      _ExtentY        =   503
      _StockProps     =   14
      Caption         =   "LIST OF CONTRACTOR"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColorLight=   16711680
      GradientColorDark=   16711680
   End
   Begin VB.Label lblLINENO 
      Caption         =   "Label8"
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   5280
      Width           =   3030
   End
End
Attribute VB_Name = "frmCSMSContractorMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim theRo                                              As String
Dim thestatus                                          As String
Dim thecode                                            As String

Sub CheckIfReallyFinish(vRONO As String)
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT DONE FROM CSMS_RO_DET WHERE REP_OR = '" & vRONO & "' AND (DONE IS NULL OR DONE <> 'Y') AND LIVIL = '1'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        gconDMIS.Execute ("UPDATE CSMS_REPAIRORDER SET STATUS = 'Working' WHERE RO_NO = '" & vRONO & "'")
    Else
        gconDMIS.Execute ("UPDATE CSMS_REPAIRORDER SET STATUS = 'Finish Job' WHERE RO_NO = '" & vRONO & "'")
    End If

    Set RSTMP = Nothing
End Sub

Sub LoadContractor()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim arnie                                          As ListItem

    SQL = "SELECT * FROM CSMS_ContractorMonitoring"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    listContractor.ListItems.Clear

    Do While Not RS.EOF
        Set arnie = listContractor.ListItems.Add(, , RS!Code)
        arnie.SubItems(1) = Null2String(RS!Contractorname)
        arnie.SubItems(2) = Null2String(Left(RS!assignedro, 10))
        arnie.SubItems(3) = Null2String(RS!Status)
        arnie.SubItems(4) = Null2String(Mid(RS!assignedro, 12, Len(RS!assignedro) - 11))
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Sub loadinfo()
    Dim RS                                             As New ADODB.Recordset
    Dim SQL                                            As String

    SQL = "SELECT Plate_no, Model, Niym, Rep_or From CSMS_Repor Where Rep_or = '" & Left(theRo, 10) & "'"
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not (RS.BOF And RS.EOF) Then
        txtRO.Text = Left(theRo, 10)
        txtCUST.Text = Null2String(RS!NIYM)
        txtplate.Text = Null2String(RS!PLATE_NO)
        txtModel.Text = Null2String(RS!MODEL)
    Else
        txtRO.Text = ""
        txtCUST.Text = ""
        txtplate.Text = ""
        txtModel.Text = ""
    End If

    Set RS = Nothing
End Sub

Sub loadcon()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "SELECT * FROM CSMS_Contractor where Code = '" & thecode & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    txtContractorName.Text = Null2String(RS!Firstname) + "," + Null2String(RS!lastname)
    txtAddress.Text = Null2String(RS!CompanyName)
    txtcompany.Text = Null2String(RS!Address)
End Sub

Private Sub cmdIn_Click()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim NewStatus                                      As String
    Dim ans                                            As String

    NewStatus = "Working"

    ans = MsgBox("Are you sure do you want to Clock In?", vbQuestion + vbYesNo)

    If ans = vbYes Then
        'SQL = "UPDATE CSMS_ContractorMonitoring Set Status = '" & NewStatus & "' Where Assignedro = '" & theRo & "'"
        gconDMIS.Execute ("UPDATE CSMS_ContractorMonitoring Set Status = '" & NewStatus & "' Where Assignedro = '" & theRo & "'")

        gconDMIS.Execute "UPDATE CSMS_RepairOrder Set Status = '" & NewStatus & "' Where  RO_NO = '" & Left(theRo, 10) & "'"
        gconDMIS.Execute ("UPDATE CSMS_RO_DET SET STATUS = 'W',DONE = 'W' WHERE LINE_NO = '" & lblLINENO.Caption & "' AND REP_OR = '" & Left(theRo, 10) & "'")

        LogAudit "A", "CONTACTOR MONITORING LOG IN", "RO" & theRo

        LoadContractor
        picin.Visible = False
    End If
End Sub

Private Sub cmdOut_Click()
    If cboOUT.Text = "" Then
        MsgBox "Choose a Reason to Clock Out", vbInformation, "Contractor"
        cboOUT.SetFocus
        Exit Sub
    End If
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim NewStatus                                      As String
    Dim ans                                            As String
    Dim vlineNo                                        As Integer

    NewStatus = "Finish Job"

    ans = MsgBox("Are you sure do you want to Clock out?", vbQuestion + vbYesNo)

    If ans = vbYes Then
        'SQL = "UPDATE CSMS_ContractorMonitoring Set Status = '" & NewStatus & "' Where Assignedro = '" & theRo & "'"
        gconDMIS.Execute ("UPDATE CSMS_ContractorMonitoring Set Status = 'Available' Where Assignedro = '" & theRo & "'")

        If cboOUT.Text = "Finish Job" Then
            gconDMIS.Execute "UPDATE CSMS_RepairOrder Set Status = '" & NewStatus & "' Where RO_NO = '" & Left(theRo, 10) & "'"
            gconDMIS.Execute ("UPDATE CSMS_RO_DET SET STATUS = 'Y',DONE = 'Y' WHERE LINE_NO = '" & lblLINENO.Caption & "' AND REP_OR = '" & Left(theRo, 10) & "'")
        Else
            gconDMIS.Execute "UPDATE CSMS_RepairOrder Set Status = 'Idle Time',TECH2 = 'Job Assigned to Other.' Where RO_NO = '" & Left(theRo, 10) & "'"
            gconDMIS.Execute ("UPDATE CSMS_RO_DET SET STATUS = 'I',DONE = 'N',TECHCODE = NULL,TECHNICIAN = NULL WHERE LINE_NO = '" & lblLINENO.Caption & "' AND REP_OR = '" & Left(theRo, 10) & "'")
        End If
        gconDMIS.Execute "UPDATE CSMS_Contractormonitoring Set Assignedro = Null, Status = 'Available' WHERE CODE = '" & thecode & "'"

        'Call CheckIfReallyFinish(Left(theRo, 10))
        LogAudit "A", "CONTACTOR MONITORING LOG OUT", "RO" & theRo
        LoadContractor
        picin.Visible = False
    End If
End Sub

Private Sub cmdOutCancel_Click()
    picin.Visible = False
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    LoadContractor
    TrapNoRO.Visible = False

    picin.Visible = False

    cmdIn.Visible = False
    cmdOut.Visible = False
    'listContractor_DblClick
End Sub

Private Sub listContractor_Click()
    On Error Resume Next
    'theRo = Left(listContractor.SelectedItem.SubItems(2), 10)
    'theRo = listContractor.SelectedItem.SubItems(2)

    lblLINENO.Caption = listContractor.SelectedItem.SubItems(4)
    theRo = listContractor.SelectedItem.SubItems(2) & "-" & lblLINENO.Caption
    Call loadinfo
    If theRo = "" Then
        TrapNoRO.Visible = True
        picin.Visible = False
        Exit Sub
    End If
End Sub

Private Sub listContractor_DblClick()
    On Error Resume Next
    theRo = listContractor.SelectedItem.SubItems(2) & "-" & lblLINENO.Caption
    thestatus = listContractor.SelectedItem.SubItems(3)
    thecode = listContractor.ListItems(listContractor.SelectedItem.INDEX).Text

    If thestatus = "" Then
        picin.Visible = False
    End If

    If theRo = "" Then
        TrapNoRO.Visible = True
        picin.Visible = False
        Exit Sub
    End If
    TrapNoRO.Visible = False
    loadinfo

    If thestatus = "Assigned" Then
        picin.Visible = True
        cboOUT.Visible = False
        cmdIn.Visible = True
        cmdOut.Visible = False
    End If

    If thestatus = "Working" Then
        picin.Visible = True
        cmdIn.Visible = False
        cboOUT.Visible = True
        cmdOut.Visible = True
    End If

    loadcon
End Sub

Private Sub Timer3_Timer()
    If Label19.ForeColor = &HC0& Then
        Label19.ForeColor = &HC0C0&
    Else
        Label19.ForeColor = &HC0&
    End If
End Sub


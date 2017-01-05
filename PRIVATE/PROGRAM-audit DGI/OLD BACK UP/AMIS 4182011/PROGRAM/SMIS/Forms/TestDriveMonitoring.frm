VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmFile_TestDriveMonitoring 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Drive Monitoring"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12270
   Icon            =   "TestDriveMonitoring.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   12270
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Test Drive Filtering"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   12255
      Begin VB.CheckBox chkall 
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4020
         TabIndex        =   23
         Top             =   1260
         Width           =   1395
      End
      Begin VB.CheckBox ChkDisApp 
         Caption         =   "DisApprove"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2580
         TabIndex        =   12
         Top             =   1230
         Width           =   1395
      End
      Begin VB.CheckBox ChkApp 
         Caption         =   "Approve"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1470
         TabIndex        =   11
         Top             =   1230
         Width           =   1125
      End
      Begin VB.Frame Frame3 
         Caption         =   "Option"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   9000
         TabIndex        =   10
         Top             =   150
         Width           =   3165
         Begin wizButton.cmd CmdView 
            Height          =   405
            Left            =   1680
            TabIndex        =   14
            Top             =   270
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   714
            TX              =   "View"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "TestDriveMonitoring.frx":08CA
         End
         Begin wizButton.cmd cmdPrint 
            Height          =   405
            Left            =   1680
            TabIndex        =   15
            Top             =   270
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   714
            TX              =   "Print"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FOCUSR          =   -1  'True
            MPTR            =   0
            MICON           =   "TestDriveMonitoring.frx":08E6
         End
      End
      Begin VB.CheckBox chkSA 
         Height          =   360
         Left            =   4800
         TabIndex        =   8
         Top             =   390
         Width           =   225
      End
      Begin VB.CheckBox ChkDate 
         Height          =   360
         Left            =   4830
         TabIndex        =   7
         Top             =   840
         Width           =   225
      End
      Begin VB.ComboBox Cbosa 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1470
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   390
         Width           =   3285
      End
      Begin MSComCtl2.DTPicker DTdate 
         Height          =   345
         Left            =   1470
         TabIndex        =   5
         Top             =   780
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54591489
         CurrentDate     =   39462
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   1230
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Agent Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   210
         TabIndex        =   3
         Top             =   390
         Width           =   1215
      End
   End
   Begin wizButton.cmd cmdClose 
      Height          =   405
      Left            =   10920
      TabIndex        =   9
      Top             =   7530
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "TestDriveMonitoring.frx":0902
   End
   Begin wizButton.cmd cmdforApp 
      Height          =   405
      Left            =   8850
      TabIndex        =   18
      Top             =   7530
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
      TX              =   "View All For Approval"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "TestDriveMonitoring.frx":091E
   End
   Begin VB.PictureBox picforapproval 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   5835
      Left            =   2880
      ScaleHeight     =   5805
      ScaleWidth      =   7485
      TabIndex        =   16
      Top             =   960
      Width           =   7515
      Begin VB.TextBox txtsearchSa 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1440
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         Height          =   255
         Left            =   7140
         TabIndex        =   19
         Top             =   60
         Width           =   315
      End
      Begin MSComctlLib.ListView listforapp 
         Height          =   4845
         Left            =   60
         TabIndex        =   22
         Top             =   900
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   8546
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Vehicle Model"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "CS no"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Color"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "SAE"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ProspectID"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SAE Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1215
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   435
         Left            =   -30
         TabIndex        =   17
         Top             =   -30
         Width           =   7515
         _Version        =   655364
         _ExtentX        =   13256
         _ExtentY        =   767
         _StockProps     =   14
         Caption         =   ":::Test Drive for approval:::"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         Alignment       =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Test Drive List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5685
      Left            =   30
      TabIndex        =   0
      Top             =   1830
      Width           =   12225
      Begin MSComctlLib.ListView ListTestDrive 
         Height          =   5295
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   9340
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Model"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Color"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Engine No"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "CS No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Time"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Client Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Sa Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Notes"
            Object.Width           =   8819
         EndProperty
      End
   End
End
Attribute VB_Name = "frmFile_TestDriveMonitoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub LoadTestdrive()
    Dim RS                                                            As New ADODB.Recordset
    Dim SQL                                                           As String
    Dim Item                                                          As ListItem
    Dim cnt                                                           As Integer


    If chkALL.Value = 1 Then
        SQL = "SELECT * FROM CRIS_MrrInv "
    Else
        SQL = "SELECT * FROM CRIS_MrrInv where saname='" & cboSA.Text & "' and datereceived='" & DTdate.Value & "'"
    End If
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    ListTestDrive(0).ListItems.Clear

    cnt = 0

    If RS.EOF And RS.BOF Then
        MsgBox "No Item found...", vbInformation, "information"
        Exit Sub
    End If

    Do While Not RS.EOF
        cnt = cnt + 1
        Set Item = ListTestDrive(0).ListItems.Add(, , cnt)
        Item.SubItems(1) = Null2String(RS!DESCRIPT)
        Item.SubItems(2) = Null2String(RS!Model)
        Item.SubItems(3) = Null2String(RS!Color)
        Item.SubItems(4) = Null2String(RS!ENGINENUMBER)
        Item.SubItems(5) = Null2String(RS!IGNKEYNO)
        Item.SubItems(6) = Null2String(RS!datereceived)
        Item.SubItems(7) = Null2String(RS!Time)
        Item.SubItems(8) = Null2String(RS!Clientname)
        Item.SubItems(9) = Null2String(RS!saname)
        Item.SubItems(10) = Null2String(RS!STATUS)
        Item.SubItems(11) = Null2String(RS!Notes)
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Sub FillSAE()
    Dim RsSAE                                                         As New ADODB.Recordset
    Dim SQL                                                           As String
    Dim MIDDLE                                                        As String
    Dim X                                                             As String

    SQL = "SELECT lname,fname,middle FROM Smis_SalesTeam"


    Set RsSAE = New ADODB.Recordset
    Set RsSAE = gconDMIS.Execute(SQL)

    cboSA.Clear

    Do While Not RsSAE.EOF
        MIDDLE = Null2String(RsSAE!MIDDLE)
        X = Mid(MIDDLE, 1, 1)
        cboSA.AddItem Null2String(RsSAE!lname) + "," + Null2String(RsSAE!fname) + "." + X

        RsSAE.MoveNext
    Loop
    Set RsSAE = Nothing

End Sub

Sub LoadAllforApproval()
    Dim RS                                                            As New ADODB.Recordset
    Dim SQL                                                           As String
    Dim cnt                                                           As Integer
    Dim Item                                                          As ListItem
    Dim xstatus                                                       As String
    xstatus = "For Approval"

    SQL = "SELECT prospectID,vehiclemodel,vehiclecode,color,status,SAE from CRIS_TestdriveSchedules where Status='" & xstatus & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)
    LogAudit "V", "TEST DRIVE MONITORING FOR APPROVAL" & " SA:" & cboSA
    listforapp.ListItems.Clear

    cnt = 0

    Do While Not RS.EOF
        cnt = cnt + 1
        Set Item = listforapp.ListItems.Add(, , cnt)
        Item.SubItems(1) = Null2String(RS!vehiclemodel)
        Item.SubItems(2) = Null2String(RS!vehiclecode)
        Item.SubItems(3) = Null2String(RS!Color)
        Item.SubItems(4) = Null2String(RS!STATUS)
        Item.SubItems(5) = Null2String(RS!SAE)
        Item.SubItems(6) = Null2String(RS!PROSPECTID)
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Sub searchMe()
    Dim RS                                                            As New ADODB.Recordset
    Dim SQL                                                           As String
    Dim cnt                                                           As Integer
    Dim Item                                                          As ListItem
    Dim Keyword                                                       As String

    Keyword = Trim(txtsearchSa.Text)

    If Keyword = "" Then
        LoadAllforApproval
        Exit Sub
    End If

    SQL = "SELECT prospectID,vehiclemodel,vehiclecode,color,status,SAE from CRIS_TestdriveSchedules where"

    If Len(Keyword) = 0 Then Exit Sub

    SQL = SQL & " SAE LIKE  '" & Keyword & "%'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    listforapp.ListItems.Clear

    cnt = 0

    Do While Not RS.EOF
        cnt = cnt + 1
        Set Item = listforapp.ListItems.Add(, , cnt)
        Item.SubItems(1) = Null2String(RS!vehiclemodel)
        Item.SubItems(2) = Null2String(RS!vehiclecode)
        Item.SubItems(3) = Null2String(RS!Color)
        Item.SubItems(4) = Null2String(RS!STATUS)
        Item.SubItems(5) = Null2String(RS!SAE)
        Item.SubItems(6) = Null2String(RS!PROSPECTID)
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Private Sub ChkDate_Click()
    If ChkDate.Value = 1 Then
        DTdate.Enabled = False
    Else
        DTdate.Enabled = True
    End If
End Sub

Private Sub chkSA_Click()

    If chkSA.Value = 1 Then
        cboSA.Text = ""
        cboSA.Enabled = False
    Else
        cboSA.Enabled = True
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdforApp_Click()
    picforapproval.Visible = True
    LoadAllforApproval
End Sub

Private Sub CmdView_Click()
    LoadTestdrive
End Sub

Private Sub Command1_Click()
    picforapproval.Visible = False
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    chkALL.Value = 1
    LoadTestdrive
    FillSAE
    picforapproval.Visible = False
    txtsearchSa.Text = ""
End Sub

Private Sub txtsearchSa_Change()
    searchMe
End Sub


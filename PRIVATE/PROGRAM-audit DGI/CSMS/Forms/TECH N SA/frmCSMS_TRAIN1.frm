VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCSMS_TECHTRAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Technician Training/Seminar Information"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9180
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_TRAIN1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   9180
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   9165
      TabIndex        =   18
      Top             =   2940
      Width           =   9165
      Begin MSComctlLib.ListView lsvLIST 
         Height          =   1575
         Left            =   60
         TabIndex        =   19
         Top             =   30
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Training"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Month-Year"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Place"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Sponsor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.PictureBox PicInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3195
      Left            =   60
      ScaleHeight     =   3195
      ScaleWidth      =   9165
      TabIndex        =   5
      Top             =   -270
      Width           =   9165
      Begin VB.TextBox txtMonYear 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1920
         MaxLength       =   10
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   990
         Width           =   2925
      End
      Begin VB.TextBox txtPlace 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   585
         Left            =   1920
         MaxLength       =   35
         TabIndex        =   8
         Top             =   1440
         Width           =   7080
      End
      Begin VB.TextBox txtTraining 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   1920
         MaxLength       =   75
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   360
         Width           =   7095
      End
      Begin VB.TextBox txtSponsor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   945
         Left            =   1920
         MaxLength       =   75
         TabIndex        =   6
         Top             =   2100
         Width           =   7095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Month-Year"
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
         Height          =   240
         Left            =   825
         TabIndex        =   17
         Top             =   1020
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Place"
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
         Height          =   240
         Left            =   1335
         TabIndex        =   16
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Training Title"
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
         Height          =   240
         Left            =   735
         TabIndex        =   15
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sponsor"
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
         Height          =   240
         Index           =   0
         Left            =   1110
         TabIndex        =   14
         Top             =   2190
         Width           =   720
      End
      Begin VB.Label labID 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         Left            =   7770
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblID 
         BackColor       =   &H000000FF&
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
         Left            =   5490
         TabIndex        =   12
         Top             =   1050
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblEMPLEVEL 
         BackColor       =   &H000000FF&
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
         Left            =   150
         TabIndex        =   11
         Top             =   2670
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblEMPNO 
         BackColor       =   &H000000FF&
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
         Left            =   180
         TabIndex        =   10
         Top             =   1950
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.PictureBox picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   6300
      ScaleHeight     =   825
      ScaleWidth      =   2925
      TabIndex        =   0
      Top             =   4590
      Width           =   2925
      Begin VB.CommandButton cmdexit 
         Caption         =   "E&xit"
         Height          =   765
         Left            =   2070
         MouseIcon       =   "frmCSMS_TRAIN1.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TRAIN1.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmddelete 
         Caption         =   "&Delete"
         Height          =   765
         Left            =   1380
         MouseIcon       =   "frmCSMS_TRAIN1.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TRAIN1.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Edit"
         Height          =   765
         Left            =   690
         MouseIcon       =   "frmCSMS_TRAIN1.frx":11FF
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TRAIN1.frx":1351
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "&Add"
         Height          =   765
         Left            =   0
         MouseIcon       =   "frmCSMS_TRAIN1.frx":16AD
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TRAIN1.frx":17FF
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picture2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   7680
      ScaleHeight     =   825
      ScaleWidth      =   1605
      TabIndex        =   20
      Top             =   4560
      Width           =   1605
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   765
         Left            =   690
         MouseIcon       =   "frmCSMS_TRAIN1.frx":1B12
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TRAIN1.frx":1C64
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cancel Entry"
         Top             =   60
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   765
         Left            =   0
         MouseIcon       =   "frmCSMS_TRAIN1.frx":1FA2
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TRAIN1.frx":20F4
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Save Entry"
         Top             =   60
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmCSMS_TECHTRAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ADD_EDIT                                           As String

Function PutValue()
    lblEMPLEVEL.Caption = frmCSMS_TECH.lblEMPTYPE
    lblEMPNO.Caption = frmCSMS_TECH.txtENO
End Function

Public Function LimitChar(ByVal alpha As String, ByVal k As Integer)
    If InStr(alpha, Chr(k)) > 0 Or k = 8 Then
        LimitChar = k
    Else
        LimitChar = 0
    End If
End Function

Sub DisAblePics(COND As Boolean)
    Picture1.Visible = Not COND
    Picture2.Visible = COND
    picList.Enabled = Not COND
    picInfo.Enabled = COND
End Sub

Sub initMemvars()
    txtTraining.Text = ""
    txtMonYear.Text = ""
    txtPlace.Text = ""
    txtSponsor.Text = ""
End Sub

Sub FillTheList()
    Dim RSTMP                                          As New ADODB.Recordset
    Dim ITEM                                           As ListItem

    'Set rsTmp = gconDMIS.Execute("Select * from CSMS_SATRAIN Where Empno = '" & frmCSMS_SA.txtEmpNo.Text & "' Order by ID ASC")
    Set RSTMP = gconDMIS.Execute("Select * from HRMS_TRAINING Where Empno = '" & lblEMPNO.Caption & "' Order by ID ASC")
    lsvLIST.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvLIST.ListItems.Add(, , Null2String(RSTMP!training))
            ITEM.SubItems(1) = Null2String(RSTMP!monyear)
            ITEM.SubItems(2) = Null2String(RSTMP!place)
            ITEM.SubItems(3) = Null2String(RSTMP!sponsor)
            ITEM.SubItems(4) = Null2String(RSTMP!ID)

            RSTMP.MoveNext
        Loop
    Else
        ShowNoRecord
        Call cmdAdd_Click
    End If

    If Not lsvLIST.ListItems.Count = 0 Then Call lsvLIST_Click

    Set RSTMP = Nothing
End Sub

Private Sub cmdAdd_Click()
    On Error Resume Next
    If Function_Access(LOGID, "ACESS_ADD", "TECHNICIAN") = False Then Exit Sub
    ADD_EDIT = "ADD"

    initMemvars
    Call DisAblePics(True)

    txtTraining.SetFocus
End Sub

Private Sub cmdCancel_Click()
    If Not lsvLIST.ListItems.Count = 0 Then
        lsvLIST_Click
        Call DisAblePics(False)
    Else
        ShowNoRecord
        Call cmdAdd_Click
    End If
End Sub

Private Sub cmdDelete_Click()
    If Not txtTraining.Text = "" Then
        If Function_Access(LOGID, "ACESS_DELETE", "TECHNICIAN") = False Then Exit Sub

        If MsgBox("Delete " & txtTraining.Text & " Training Plan", vbQuestion + vbYesNo, "Are You Sure") = vbYes Then
            SQL_STATEMENT = "Delete From HRMS_TRAINING Where Empno = '" & lblEMPNO & _
                            "' And ID = " & lblID.Caption & ""
            gconDMIS.Execute SQL_STATEMENT

            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("XX", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(lblEMPNO), "EMPNO", "HRMS_EMPINFO"), "", "TITLE: " & txtTraining, "", labid)
            'NEW LOG AUDIT-----------------------------------------------------

            ShowDeletedMsg
            FillTheList
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    If Not txtTraining.Text = "" Then
        If Function_Access(LOGID, "ACESS_EDIT", "TECHNICIAN") = False Then Exit Sub

        ADD_EDIT = "EDIT"

        Call DisAblePics(True)
        txtTraining.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim vTRAIN                                         As String
    Dim vDATE                                          As String
    Dim vPLACE                                         As String
    Dim vSPON                                          As String
    Dim vDET                                           As String

    If txtTraining.Text = "" Then
        MsgBox "Training Title Cannot be Blank", vbInformation, "Training Plan"
        txtTraining.SetFocus
        Exit Sub
    End If
    If txtMonYear.Text = "" Then
        MsgBox "Date Taken Cannot be Blank", vbInformation, "Training Plan"
        txtMonYear.SetFocus
        Exit Sub
    Else
        If IsDate(txtMonYear) = False Then
            MsgBox "Invalid date Format", vbExclamation, "CSMS"
            txtMonYear.SetFocus
            Exit Sub
        End If
    End If
    If txtPlace.Text = "" Then
        MsgBox "Place Name Cannot be Blank", vbInformation, "Training Plan"
        txtPlace.SetFocus
        Exit Sub
    End If
    If txtSponsor.Text = "" Then
        MsgBox "Sponsor Cannot be Blank", vbInformation, "Training Plan"
        txtSponsor.SetFocus
        Exit Sub
    End If


    vTRAIN = N2Str2Null(txtTraining.Text)
    vDATE = N2Str2Null(txtMonYear.Text)
    vPLACE = N2Str2Null(txtPlace.Text)
    vSPON = N2Str2Null(txtSponsor.Text)
    vDET = N2Str2Null(txtMonYear.Text)

    If ADD_EDIT = "ADD" Then
        'gconDMIS.Execute ("Insert Into CSMS_SATRAIN (Empno,Training,Deyt,Place,Sponsor,Details) VALUES('" & frmCSMS_SA.txtEmpNo.Text & _
         "'," & vTRAIN & "," & vDATE & "," & vPLACE & "," & vSPON & "," & vDET & ")")

        SQL_STATEMENT = "INSERT INTO HRMS_TRAINING (EMPLEVEL,EMPNO,TRAINING,MONYEAR,PLACE,SPONSOR,USERCODE,LASTUPDATE) VALUES " & _
                        "(" & N2Str2Null(lblEMPLEVEL) & "," & N2Str2Null(lblEMPNO.Caption) & "," & vTRAIN & "," & vDET & "," & vPLACE & "," & vSPON & "," & N2Str2Null(LOGCODE) & "," & N2Str2Null(LOGDATE) & ")"
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("AA", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(lblEMPNO), "EMPNO", "HRMS_EMPINFO"), "", "TRAINING: " & txtTraining, "", FindTransactionID(lblEMPNO, "EMPNO", "HRMS_TRAINING", "DETAILS", N2Str2Null(txtTraining), "TRAINING"))
        'NEW LOG AUDIT-----------------------------------------------------

        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "UPDATE HRMS_TRAINING SET Training = " & vTRAIN & _
                        ",MONYEAR = " & vDATE & _
                        ",PLACE = " & vPLACE & _
                        ",SPONSOR = " & vSPON & _
                        ",USERCODE = " & N2Str2Null(LOGCODE) & _
                        ",LASTUPDATE = " & N2Str2Null(LOGDATE) & _
                      " Where Empno = '" & lblEMPLEVEL & "' And ID = " & lblID.Caption & ""
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("EE", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(lblEMPNO), "EMPNO", "HRMS_EMPINFO"), "", "TRAIINING: " & txtTraining, "", labid)
        'NEW LOG AUDIT-----------------------------------------------------

        ShowSuccessFullyUpdated
    End If

    Call DisAblePics(False)
    FillTheList
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    PutValue

    FillTheList
End Sub

Private Sub lsvLIST_Click()
    Dim Index                                          As Double

    If Not lsvLIST.ListItems.Count = 0 Then
        With lsvLIST
            Index = .SelectedItem.Index

            txtTraining.Text = .ListItems(Index).Text
            txtMonYear.Text = .ListItems(Index).SubItems(1)
            txtPlace.Text = .ListItems(Index).SubItems(2)
            txtSponsor.Text = .ListItems(Index).SubItems(3)
            lblID.Caption = .ListItems(Index).SubItems(4)
        End With
    End If
End Sub

Private Sub lsvLIST_DblClick()
    If lsvLIST.ListItems.Count = 0 Then Exit Sub

    Dim Index                                          As Integer
    Index = lsvLIST.SelectedItem.Index

    Call cmdEdit_Click
End Sub

Private Sub txtMonYear_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890/", KeyAscii)
    End If
End Sub


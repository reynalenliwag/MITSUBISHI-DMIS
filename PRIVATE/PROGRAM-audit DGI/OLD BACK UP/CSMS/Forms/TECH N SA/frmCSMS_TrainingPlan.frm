VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCSMS_SATrainingPlan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Adviser Training Plans"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_TrainingPlan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9840
   Begin VB.PictureBox picInfo 
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
      Height          =   3075
      Left            =   90
      ScaleHeight     =   3075
      ScaleWidth      =   9705
      TabIndex        =   15
      Top             =   60
      Width           =   9705
      Begin VB.TextBox txtNameOfTraining 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   885
         Left            =   2340
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   810
         Width           =   7305
      End
      Begin VB.TextBox txtDateComp 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2340
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2640
         Width           =   3465
      End
      Begin VB.TextBox txtDevType 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2340
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   2235
         Width           =   3465
      End
      Begin VB.TextBox txtDateSched 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2340
         MaxLength       =   10
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1785
         Width           =   3465
      End
      Begin VB.TextBox txtDesPer 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   2340
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   90
         Width           =   7335
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
         Height          =   240
         Left            =   8700
         TabIndex        =   24
         Top             =   2280
         Visible         =   0   'False
         Width           =   735
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
         Height          =   240
         Left            =   8700
         TabIndex        =   23
         Top             =   1920
         Visible         =   0   'False
         Width           =   735
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
         Height          =   240
         Left            =   8700
         TabIndex        =   22
         Top             =   2670
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Training /Development to fulfill Desired Performance"
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
         Height          =   975
         Left            =   150
         TabIndex        =   21
         Top             =   810
         Width           =   2115
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
         Left            =   3390
         TabIndex        =   20
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Completed"
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
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Desired Performance /Competency"
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
         Height          =   1005
         Left            =   150
         TabIndex        =   18
         Top             =   120
         Width           =   2115
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Training/Dev. Type"
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
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   2310
         Width           =   2715
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Approx. Date to be Sched."
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
         Height          =   465
         Left            =   150
         TabIndex        =   16
         Top             =   1740
         Width           =   2115
      End
   End
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
      Height          =   1905
      Left            =   90
      ScaleHeight     =   1905
      ScaleWidth      =   9705
      TabIndex        =   14
      Top             =   3060
      Width           =   9705
      Begin MSComctlLib.ListView lsvLIST 
         Height          =   1785
         Left            =   30
         TabIndex        =   5
         Top             =   90
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   3149
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Desired Performance"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name Of Training"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date To Take"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date Finish"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   0
         EndProperty
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
      Height          =   885
      Left            =   6840
      ScaleHeight     =   885
      ScaleWidth      =   3105
      TabIndex        =   13
      Top             =   5010
      Width           =   3105
      Begin VB.CommandButton cmdexit 
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
         Left            =   2160
         MouseIcon       =   "frmCSMS_TrainingPlan.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TrainingPlan.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmddelete 
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
         Left            =   1440
         MouseIcon       =   "frmCSMS_TrainingPlan.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TrainingPlan.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Delete Selected Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdedit 
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
         Left            =   720
         MouseIcon       =   "frmCSMS_TrainingPlan.frx":11FF
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TrainingPlan.frx":1351
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdadd 
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
         Left            =   0
         MouseIcon       =   "frmCSMS_TrainingPlan.frx":16AD
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TrainingPlan.frx":17FF
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   735
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
      Left            =   8310
      ScaleHeight     =   825
      ScaleWidth      =   1605
      TabIndex        =   12
      Top             =   4980
      Width           =   1605
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
         MouseIcon       =   "frmCSMS_TrainingPlan.frx":1B12
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TrainingPlan.frx":1C64
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancel Entry"
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
         Left            =   0
         MouseIcon       =   "frmCSMS_TrainingPlan.frx":1FA2
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_TrainingPlan.frx":20F4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmCSMS_SATrainingPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ADD_EDIT                                           As String

Function PutValue()
    lblEMPLEVEL.Caption = frmCSMS_SA.lblEMPTYPE
    lblEMPNO.Caption = frmCSMS_SA.txtENO
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
    txtDateComp.Text = ""
    txtDateSched.Text = ""
    txtDesPer.Text = ""
    txtDevType.Text = ""
    txtNameOfTraining.Text = ""
End Sub

Sub FillTheList()
    Dim RSTMP                                          As New ADODB.Recordset
    Dim ITEM                                           As ListItem

    Set RSTMP = gconDMIS.Execute("Select * from HRMS_TRAININGPLAN Where Empno = '" & lblEMPNO.Caption & "' Order by ID ASC")
    lsvLIST.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvLIST.ListItems.Add(, , Null2String(RSTMP!DESPER))
            ITEM.SubItems(1) = Null2String(RSTMP!NAMEOFTRAINING)
            ITEM.SubItems(2) = Null2String(RSTMP!DATESCHED)
            ITEM.SubItems(3) = Null2String(RSTMP!DEVTYPE)
            ITEM.SubItems(4) = Null2String(RSTMP!DateComp)
            ITEM.SubItems(5) = Null2String(RSTMP!ID)

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
    If Function_Access(LOGID, "ACESS_ADD", "SERVICE ADVISOR") = False Then Exit Sub

    ADD_EDIT = "ADD"

    initMemvars
    DisAblePics True

    txtDesPer.SetFocus
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
    If Not txtDesPer.Text = "" Then
        If Function_Access(LOGID, "ACESS_DELETE", "SERVICE ADVISOR") = False Then Exit Sub

        If MsgBox("Delete " & txtDesPer.Text & " Training Plan", vbQuestion + vbYesNo, "Are You Sure") = vbYes Then
            SQL_STATEMENT = "Delete From HRMS_TRAININGPLAN Where Empno = '" & lblEMPNO & _
                            "' And ID = " & lblID.Caption & ""
            gconDMIS.Execute SQL_STATEMENT
            'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("XX", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(lblEMPNO), "EMPNO", "HRMS_EMPINFO"), "", "TITLE: " & txtDesPer, "", labid)
            'NEW LOG AUDIT-----------------------------------------------------

            ShowDeletedMsg
            FillTheList
        End If
    End If
End Sub

Private Sub cmdEdit_Click()
    If Not txtDesPer.Text = "" Then
        If Function_Access(LOGID, "ACESS_EDIT", "SERVICE ADVISOR") = False Then Exit Sub

        ADD_EDIT = "EDIT"

        Call DisAblePics(True)
        txtDesPer.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim vPER                                           As String
    Dim vTRAIN                                         As String
    Dim vDATE                                          As String
    Dim VTYPE                                          As String
    Dim vDATEC                                         As String

    If txtDesPer.Text = "" Then
        ShowIsRequiredMsg "Desired Performance Cannot be Blank"
        txtDesPer.SetFocus
        Exit Sub
    End If
    If txtNameOfTraining.Text = "" Then
        ShowIsRequiredMsg "Name of Training Cannot be Blank"
        txtNameOfTraining.SetFocus
        Exit Sub
    End If
    If txtDateSched.Text = "" Then
        ShowIsRequiredMsg "Date Schedule Cannot be Blank"
        txtDateSched.SetFocus
        Exit Sub
    Else
        If IsDate(txtDateSched) = False Then
            MsgBox "Invalid Date Format", vbExclamation, "CSMS"
            txtDateSched.SetFocus
            Exit Sub
        End If
    End If
    If txtDevType.Text = "" Then
        MsgBox "Type Cannot be Blank", vbInformation, "Training Plan"
        txtDevType.SetFocus
        Exit Sub
    End If

    vPER = N2Str2Null(txtDesPer.Text)
    vTRAIN = N2Str2Null(txtNameOfTraining.Text)
    vDATE = N2Str2Null(txtDateSched.Text)
    VTYPE = N2Str2Null(txtDevType.Text)
    vDATEC = N2Str2Null(txtDateComp.Text)

    If ADD_EDIT = "ADD" Then
        SQL_STATEMENT = "Insert Into HRMS_TRAININGPLAN (Empno,DesPer,NAMEOFTRAINING,DateSCHED,DEVType,DateComp,USERCODE,LASTUPDATE) VALUES " & _
                        "(" & N2Str2Null(lblEMPNO.Caption) & _
                        "," & vPER & _
                        "," & vTRAIN & _
                        "," & vDATE & _
                        "," & VTYPE & _
                        "," & vDATEC & _
                        "," & N2Str2Null(LOGCODE) & _
                        "," & N2Str2Null(LOGDATE) & ")"
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("AA", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(lblEMPNO), "EMPNO", "HRMS_EMPINFO"), "", "TRAINING PLAN: " & txtDesPer, "", FindTransactionID(lblEMPNO, "EMPNO", "HRMS_TRAININGPLAN", "DETAILS", N2Str2Null(txtDesPer), "DESPER"))
        'NEW LOG AUDIT-----------------------------------------------------

        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "UPDATE HRMS_TRAININGPLAN SET DesPer = " & vPER & _
                        ",NAMEOFTRAINING = " & vTRAIN & _
                        ",DATESCHED = " & vDATE & _
                        ",DEVTYPE = " & VTYPE & _
                        ",DateComp = " & vDATEC & _
                        ",USERCODE = " & N2Str2Null(LOGCODE) & _
                        ",LASTUPDATE = " & N2Str2Null(LOGDATE) & _
                      " Where Empno = '" & lblEMPNO & "' And ID = " & lblID.Caption & ""
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("EE", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(lblEMPNO), "EMPNO", "HRMS_EMPINFO"), "", "TRAINING PLAN: " & txtDesPer, "", labid)
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

            txtDesPer.Text = .ListItems(Index).Text
            txtNameOfTraining.Text = .ListItems(Index).SubItems(1)
            txtDateSched.Text = .ListItems(Index).SubItems(2)
            txtDevType.Text = .ListItems(Index).SubItems(3)
            txtDateComp.Text = .ListItems(Index).SubItems(4)
            lblID.Caption = .ListItems(Index).SubItems(5)
        End With
    End If
End Sub

Private Sub lsvLIST_DblClick()
    If lsvLIST.ListItems.Count = 0 Then Exit Sub

    Dim Index                                          As Integer
    Index = lsvLIST.SelectedItem.Index

    Call cmdEdit_Click
End Sub

Private Sub txtDateComp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890/", KeyAscii)
    End If
End Sub

Private Sub txtDateComp_LostFocus()
    'If IsDate(txtDateComp) = False Then txtDateComp = "": Exit Sub
    'txtDateComp = FormatDateTime(txtDateComp, vbShortDate)
End Sub

Private Sub txtDateSched_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890/", KeyAscii)
    End If
End Sub

Private Sub txtDateSched_LostFocus()
    'If IsDate(txtDateSched) = False Then txtDateSched = "": Exit Sub
    'txtDateSched = FormatDateTime(txtDateSched, vbShortDate)
End Sub


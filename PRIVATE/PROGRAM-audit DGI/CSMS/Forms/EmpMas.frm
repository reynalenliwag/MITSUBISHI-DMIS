VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSEmpNo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Advisor Data Entry"
   ClientHeight    =   3705
   ClientLeft      =   720
   ClientTop       =   330
   ClientWidth     =   8895
   ForeColor       =   &H00DEDFDE&
   Icon            =   "EmpMas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3705
   ScaleWidth      =   8895
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   2880
      ScaleHeight     =   945
      ScaleWidth      =   6315
      TabIndex        =   18
      Top             =   2760
      Width           =   6315
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
         Left            =   5220
         MouseIcon       =   "EmpMas.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Exit Window"
         Top             =   60
         Width           =   735
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
         Left            =   4500
         MouseIcon       =   "EmpMas.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Print this Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
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
         Left            =   3780
         MouseIcon       =   "EmpMas.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Delete Selected Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
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
         Left            =   3060
         MouseIcon       =   "EmpMas.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
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
         Left            =   2340
         MouseIcon       =   "EmpMas.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Add Record"
         Top             =   60
         Width           =   735
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
         Left            =   1620
         MouseIcon       =   "EmpMas.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Find a Record"
         Top             =   60
         Width           =   735
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
         Left            =   900
         MouseIcon       =   "EmpMas.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Move to Next Record"
         Top             =   60
         Width           =   735
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
         Left            =   180
         MouseIcon       =   "EmpMas.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7305
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   27
      Top             =   2745
      Width           =   1800
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
         Left            =   780
         MouseIcon       =   "EmpMas.frx":2D71
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":2EC3
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Cancel"
         Top             =   60
         Width           =   735
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
         Left            =   60
         MouseIcon       =   "EmpMas.frx":3201
         MousePointer    =   99  'Custom
         Picture         =   "EmpMas.frx":3353
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   2670
      TabIndex        =   5
      Top             =   30
      Width           =   6165
      Begin Crystal.CrystalReport rptSA 
         Left            =   5400
         Top             =   2160
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Service Advisor's Master List"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowGroupTree=   -1  'True
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.TextBox txtNaym 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1350
         TabIndex        =   13
         Top             =   1770
         Width           =   4695
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1350
         TabIndex        =   0
         Top             =   210
         Width           =   1065
      End
      Begin VB.TextBox txtEmpNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1350
         TabIndex        =   4
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtMiddleInt 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1350
         TabIndex        =   3
         Top             =   1380
         Width           =   615
      End
      Begin VB.TextBox txtFirstName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1350
         TabIndex        =   2
         Top             =   990
         Width           =   4695
      End
      Begin VB.TextBox txtLastName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1350
         TabIndex        =   1
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   780
         TabIndex        =   12
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Emp No"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   540
         TabIndex        =   9
         Top             =   2190
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Initial"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   8
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   270
         TabIndex        =   7
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   270
         TabIndex        =   6
         Top             =   660
         Width           =   1065
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3495
      Left            =   60
      TabIndex        =   15
      Top             =   45
      Width           =   2565
      Begin VB.TextBox textSearch 
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
         Left            =   90
         MaxLength       =   35
         TabIndex        =   16
         Top             =   150
         Width           =   2415
      End
      Begin MSComctlLib.ListView lstServiceAdvisor 
         Height          =   2895
         Left            =   60
         TabIndex        =   17
         Top             =   540
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   5106
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
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
         MouseIcon       =   "EmpMas.frx":36A3
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FULL NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   270
      TabIndex        =   11
      Top             =   420
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   600
      TabIndex        =   10
      Top             =   270
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmCSMSEmpNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmpNo                                                           As ADODB.Recordset
Dim AddorEdit                                                         As String

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "SERVICE ADVISOR") = False Then Exit Sub

    Screen.MousePointer = 11
    PrintSQLReport rptSA, CSMS_REPORT_PATH & "sa.rpt", "", CSMS_REPORT_CONNECTION, 1
    LogAudit "V", "SERVICE ADVISOR INFORMATION REPORT "
    Screen.MousePointer = 0
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "SERVICE ADVISOR") = False Then Exit Sub

    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    InitMemvars
    On Error Resume Next
    txtCODE.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "SERVICE ADVISOR") = False Then Exit Sub

    On Error GoTo Errorcode
    If Not rsEmpNo.BOF Or Not rsEmpNo.EOF Then
        MsgSpeechBox "Delete a Record? Are you Sure?"
        If MsgBoxXP("Are you sure?", "Confirm Delete", XP_YesNo, msg_Question) = True Then
            gconDMIS.Execute "delete from CSMS_vw_EmpNo where id = " & labid.Caption
            LogAudit "X", "SERVICE ADVISOR INFORMATION", "CODE/LASTNAME " & txtCODE & "/" & txtLastName
            ShowDeletedMsg
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    rsRefresh
    StoreMemVars
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub
Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "SERVICE ADVISOR") = False Then Exit Sub

    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
    txtCODE.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub
Private Sub cmdNext_Click()
    On Error Resume Next
    rsEmpNo.MoveNext
    If rsEmpNo.EOF Then
        rsEmpNo.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsEmpNo.MovePrevious
    If rsEmpNo.BOF Then
        rsEmpNo.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub
Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    If IsNull(txtCODE.Text) = True Then
        MsgSpeechBox "Employee Code must not be empty"
        On Error Resume Next
        txtCODE.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Dim rsfindDup                                             As ADODB.Recordset
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select code from CSMS_vw_EmpNo where code = '" & txtCODE.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "Employee Code already exist!"
                On Error Resume Next
                txtCODE.SetFocus
                Exit Sub
            End If
        End If
    End If
    If txtLastName.Text = "" Or txtFirstName.Text = "" Then
        MsgSpeechBox "Last Name and First Name is Required"
        On Error Resume Next
        txtLastName.SetFocus
        Exit Sub
    End If

    Dim VTXTCode, VTXTLASTNAME, VTXTFIRSTNAME                         As String
    Dim VTXTMiddleInt, VTXTNaym, VTXTEmpNo                            As String

    VTXTCode = N2Str2Null(txtCODE.Text)
    VTXTLASTNAME = N2Str2Null(UCase(txtLastName.Text))
    VTXTFIRSTNAME = N2Str2Null(UCase(txtFirstName.Text))
    VTXTMiddleInt = N2Str2Null(txtMiddleInt.Text)
    VTXTNaym = N2Str2Null(txtNaym.Text)
    VTXTEmpNo = N2Str2Null(txtEmpNo.Text)

    If AddorEdit = "ADD" Then
        If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then
            rsEmpNo.MoveLast
            labid.Caption = NumericVal(rsEmpNo!ID) + 1
        End If
        gconDMIS.Execute "Insert into CSMS_vw_EmpNo" & _
                       " (code,lastname,firstname,middleint,naym,empno)" & _
                       " values (" & VTXTCode & ", " & VTXTLASTNAME & ", " & VTXTFIRSTNAME & ", " & VTXTMiddleInt & ", " & _
                       " " & VTXTNaym & ", " & VTXTEmpNo & ")"
LogAudit "A", "SERVICE ADVISOR INFORMATION", "CODE/LASTNAME " & txtCODE & "/" & txtLastName
    Else
        gconDMIS.Execute "update CSMS_vw_EmpNo set" & _
                       " code = " & VTXTCode & "," & _
                       " lastname = " & VTXTLASTNAME & "," & _
                       " firstname = " & VTXTFIRSTNAME & "," & _
                       " middleint = " & VTXTMiddleInt & "," & _
                       " naym = " & VTXTNaym & "," & _
                       " empno = " & VTXTEmpNo & _
                       " where id = " & labid.Caption
LogAudit "E", "SERVICE ADVISOR INFORMATION", "CODE/LASTNAME " & txtCODE & "/" & txtLastName
End If
    rsRefresh
    On Error Resume Next
    rsEmpNo.Find "id =" & labid.Caption
    cmdCancel.Value = True
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    CenterMe frmMain, Me, 1
    rsRefresh
    Frame1.Enabled = False
    textSearch.Text = "":
    InitMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub InitMemvars()
    txtCODE.Text = ""
    txtLastName.Text = ""
    txtFirstName.Text = ""
    txtMiddleInt.Text = ""
    txtNaym.Text = ""
    txtEmpNo.Text = ""
End Sub
Sub StoreMemVars()
    On Error Resume Next
    If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then
        labid.Caption = rsEmpNo!ID
        txtCODE.Text = Null2String(rsEmpNo!code)
        txtLastName.Text = Null2String(rsEmpNo!lastname)
        txtFirstName.Text = Null2String(rsEmpNo!Firstname)
        txtMiddleInt.Text = Null2String(rsEmpNo!middleint)
        txtNaym.Text = Null2String(rsEmpNo!naym)
        txtEmpNo.Text = Null2String(rsEmpNo!empno)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub
Sub rsRefresh()
    Set rsEmpNo = New ADODB.Recordset
    rsEmpNo.Open "select * from CSMS_vw_EmpNo order by EMPNO asc", gconDMIS, adOpenKeyset
End Sub

Private Sub lstServiceAdvisor_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsEmpNo.Bookmark = rsFind(rsEmpNo.Clone, "EMPNO", Item.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lstServiceAdvisor_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstServiceAdvisor
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstServiceAdvisor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (textSearch.Text)
    End If
End Sub

Sub FillGrid()
    Dim rsServiceAdvisor                                              As ADODB.Recordset
    lstServiceAdvisor.Enabled = False
    lstServiceAdvisor.Sorted = False: lstServiceAdvisor.ListItems.Clear
    Set rsServiceAdvisor = New ADODB.Recordset
    Set rsServiceAdvisor = gconDMIS.Execute("select Lastname + ',' + FirstName as SANAME,EMPNO from CSMS_vw_EmpNo order by Lastname,FirstName asc")
    If Not (rsServiceAdvisor.EOF And rsServiceAdvisor.BOF) Then
        Listview_Loadval Me.lstServiceAdvisor.ListItems, rsServiceAdvisor
        lstServiceAdvisor.Refresh
        lstServiceAdvisor.Enabled = True
    End If

End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsServiceAdvisor                                              As ADODB.Recordset
    lstServiceAdvisor.Sorted = False: lstServiceAdvisor.ListItems.Clear
    lstServiceAdvisor.Enabled = False
    Set rsServiceAdvisor = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsServiceAdvisor = gconDMIS.Execute("select Lastname + ',' + FirstName as SANAME, EMPNO from CSMS_vw_EmpNo where Lastname + ',' + FirstName like'" & XXX & "%'")
    If Not (rsServiceAdvisor.EOF And rsServiceAdvisor.BOF) Then
        Listview_Loadval Me.lstServiceAdvisor.ListItems, rsServiceAdvisor
        lstServiceAdvisor.Refresh
        lstServiceAdvisor.Enabled = True
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstServiceAdvisor.Enabled = True Then
            lstServiceAdvisor.SetFocus
        End If
    End If
End Sub


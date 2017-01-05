VERSION 5.00
Begin VB.Form frmHRMSProfile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Profile & Signatories"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5850
   ForeColor       =   &H00D8E9EC&
   Icon            =   "Profile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   5850
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4245
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   26
      Top             =   5265
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
         Left            =   750
         MouseIcon       =   "Profile.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Profile.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   27
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
         Left            =   60
         MouseIcon       =   "Profile.frx":079A
         MousePointer    =   99  'Custom
         Picture         =   "Profile.frx":08EC
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picProfile 
      BorderStyle     =   0  'None
      Height          =   5205
      Left            =   30
      ScaleHeight     =   5205
      ScaleWidth      =   6255
      TabIndex        =   0
      Top             =   30
      Width           =   6255
      Begin VB.TextBox txtPreparedBy 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   2040
         TabIndex        =   11
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox txtCheckedBy 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   2040
         TabIndex        =   13
         Top             =   2310
         Width           =   3615
      End
      Begin VB.TextBox txtApprovedBy 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   2040
         TabIndex        =   15
         Top             =   2700
         Width           =   3615
      End
      Begin VB.TextBox txtNotedBy1 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   2040
         TabIndex        =   25
         Top             =   4740
         Width           =   3615
      End
      Begin VB.TextBox txtGeneralManager 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   2040
         TabIndex        =   17
         Top             =   3090
         Width           =   3615
      End
      Begin VB.TextBox txtSBManager 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   2040
         TabIndex        =   21
         Top             =   3900
         Width           =   3615
      End
      Begin VB.TextBox txtAccountNo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   2040
         TabIndex        =   19
         Top             =   3510
         Width           =   3615
      End
      Begin VB.TextBox txtCorpSec 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   2040
         TabIndex        =   23
         Top             =   4320
         Width           =   3615
      End
      Begin VB.TextBox txtCompanyTINNo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1140
         Width           =   3615
      End
      Begin VB.TextBox txtCompanyAddress 
         BackColor       =   &H00FFFFFF&
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
         Height          =   585
         Left            =   2040
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   450
         Width           =   3615
      End
      Begin VB.TextBox txtCompanyName 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   2
         Top             =   60
         Width           =   3615
      End
      Begin VB.TextBox txtCompanySSSNo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1530
         Width           =   3615
      End
      Begin VB.Label labID 
         Caption         =   "Label14"
         Height          =   165
         Left            =   3540
         TabIndex        =   3
         Top             =   210
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Prepared By"
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
         Left            =   945
         TabIndex        =   10
         Top             =   1950
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Checked By"
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
         Left            =   975
         TabIndex        =   12
         Top             =   2310
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
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
         Left            =   945
         TabIndex        =   14
         Top             =   2730
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "1st Noted By"
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
         Left            =   915
         TabIndex        =   24
         Top             =   4770
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Operations Manager"
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
         Left            =   270
         TabIndex        =   16
         Top             =   3150
         Width           =   1680
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Manager"
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
         Left            =   765
         TabIndex        =   20
         Top             =   3960
         Width           =   1185
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account No."
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
         Left            =   510
         TabIndex        =   18
         Top             =   3540
         Width           =   1440
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Secretary"
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
         Left            =   1185
         TabIndex        =   22
         Top             =   4350
         Width           =   765
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company TIN No."
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
         Left            =   495
         TabIndex        =   6
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Address"
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
         Left            =   420
         TabIndex        =   4
         Top             =   480
         Width           =   1530
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
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
         Left            =   600
         TabIndex        =   1
         Top             =   90
         Width           =   1350
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company SSS No."
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
         Left            =   420
         TabIndex        =   8
         Top             =   1560
         Width           =   1530
      End
   End
End
Attribute VB_Name = "frmHRMSProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsProfile                                                         As ADODB.Recordset

Sub InitMemvars()
    txtPreparedBy.Text = ""
    txtCheckedBy.Text = ""
    txtApprovedBy.Text = ""
    txtNotedBy1.Text = ""
    txtGeneralManager.Text = ""
    txtAccountNo.Text = ""
    txtSBManager.Text = ""
    txtCorpSec.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        labID.Caption = rsProfile!ID
        txtCompanyName.Text = Null2String(rsProfile!CompanyName)
        txtCompanyAddress.Text = Null2String(rsProfile!Companyaddress)
        txtCompanyTINNo.Text = Null2String(rsProfile!companytinno)
        txtCompanySSSNo.Text = Null2String(rsProfile!companysssno)
        txtPreparedBy.Text = Null2String(rsProfile!PREPAREDBY)
        txtCheckedBy.Text = Null2String(rsProfile!CHECKEDBY)
        txtApprovedBy.Text = Null2String(rsProfile!APPROVEDBY)
        txtNotedBy1.Text = Null2String(rsProfile!NotedBy1)
        txtGeneralManager.Text = Null2String(rsProfile!GENERALMANAGER)
        txtAccountNo.Text = Null2String(rsProfile!ACCOUNTNO)
        txtSBManager.Text = Null2String(rsProfile!BankManager)
        txtCorpSec.Text = Null2String(rsProfile!SECRETARY)
    End If
End Sub

Sub rsrefresh()
    Set rsProfile = New ADODB.Recordset
    rsProfile.Open "select * from ALL_PROFILE where modulename = 'HRMS'", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim VTXTCOMPANYNAME                                               As String
    Dim VTXTCOMPANYADDRESS                                            As String
    Dim VTXTCOMPANYTINNO                                              As String
    Dim VTXTCOMPANYSSSNO                                              As String
    Dim VTXTPREPAREDBY                                                As String
    Dim VTXTCHECKEDBY                                                 As String
    Dim VTXTAPPROVEDBY                                                As String
    Dim VTXTGENERALMANAGER                                            As String
    Dim VTXTACCOUNTNO                                                 As String
    Dim VTXTSBMANAGER                                                 As String
    Dim VTXTCORPSEC                                                   As String
    Dim VTXTNOTEDBY1                                                  As String

    If txtCompanyName = "" Then
        ShowIsRequiredMsg "Company Name"
        Exit Sub
    End If
    If txtCompanyAddress = "" Then
        ShowIsRequiredMsg "Company Address"
        Exit Sub
    End If
    If txtCompanyTINNo = "" Then
        If MsgQuestionBox("TIN No. Omitted. Continue Any way?", "Warning") = False Then
            Exit Sub
        End If
    End If
    If txtCompanySSSNo = "" Then
        If MsgQuestionBox("SSS No. Omitted. Continue Any way?", "Warning") = False Then
            Exit Sub
        End If
    End If

    VTXTCOMPANYNAME = N2Str2Null(txtCompanyName.Text)
    VTXTCOMPANYADDRESS = N2Str2Null(txtCompanyAddress.Text)
    VTXTCOMPANYTINNO = N2Str2Null(txtCompanyTINNo.Text)
    VTXTCOMPANYSSSNO = N2Str2Null(txtCompanySSSNo.Text)
    VTXTPREPAREDBY = N2Str2Null(txtPreparedBy.Text)
    VTXTCHECKEDBY = N2Str2Null(txtCheckedBy.Text)
    VTXTAPPROVEDBY = N2Str2Null(txtApprovedBy.Text)
    VTXTNOTEDBY1 = N2Str2Null(txtNotedBy1.Text)
    VTXTGENERALMANAGER = N2Str2Null(txtGeneralManager.Text)
    VTXTACCOUNTNO = N2Str2Null(txtAccountNo.Text)
    VTXTSBMANAGER = N2Str2Null(txtSBManager.Text)
    VTXTCORPSEC = N2Str2Null(txtCorpSec.Text)

    gconDMIS.Execute "UPDATE ALL_PROFILE SET" & _
                   " COMPANYNAME = " & VTXTCOMPANYNAME & "," & _
                   " COMPANYADDRESS = " & VTXTCOMPANYADDRESS & "," & _
                   " COMPANYTINNO = " & VTXTCOMPANYTINNO & "," & _
                   " COMPANYSSSNO = " & VTXTCOMPANYSSSNO & "," & _
                   " PREPAREDBY = " & VTXTPREPAREDBY & "," & _
                   " CHECKEDBY = " & VTXTCHECKEDBY & "," & _
                   " APPROVEDBY = " & VTXTAPPROVEDBY & "," & _
                   " NOTEDBY1 = " & VTXTNOTEDBY1 & "," & _
                   " GENERALMANAGER = " & VTXTGENERALMANAGER & "," & _
                   " ACCOUNTNO = " & VTXTACCOUNTNO & "," & _
                   " BANKMANAGER = " & VTXTSBMANAGER & "," & _
                   " SECRETARY = " & VTXTCORPSEC & _
                   " WHERE MODULENAME ='HRMS' "


    COMPANY_NAME = txtCompanyName.Text
    COMPANY_ADDRESS = txtCompanyAddress.Text
    COMPANY_TIN = txtCompanyTINNo.Text
    PREPARED_BY = txtPreparedBy.Text
    CHECKED_BY = txtCheckedBy.Text
    APPROVED_BY = txtApprovedBy.Text
    ACCOUNT_NO = txtAccountNo.Text
    SECRETARY = txtCorpSec.Text
    NOTED_BY = txtNotedBy1.Text

    MessagePop InfoOk, "Record Saved", "HRMS Payroll Setting Sucessfully Updated", 1000, 0
    Dim FRM                                                           As Form
    For Each FRM In Forms
        If Not (UCase(FRM.NAME) = UCase("frmMain") Or UCase(FRM.NAME) = UCase("frmMainMenu") Or UCase(FRM.NAME) = UCase(Me.NAME)) Then
            Unload FRM
        End If

    Next
    Unload Me

    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    Else
        MoveKeyPress KeyCode
    End If

End Sub

Private Sub Form_Load()
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    '    With cboCUTOFF
    '        .AddItem "1st Cut-Off"
    '        .AddItem "2nd Cut-Off"
    '    End With
    '    fillcbomonth cboMonth
    '    FillcboYear cboYear
    Call rsrefresh
    Call InitMemvars
    Call StoreMemVars
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub txtCompanyAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub


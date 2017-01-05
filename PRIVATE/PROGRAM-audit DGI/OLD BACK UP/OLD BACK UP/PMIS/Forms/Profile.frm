VERSION 5.00
Begin VB.Form frmPMISProfile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Setup"
   ClientHeight    =   6540
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   7290
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Profile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   7290
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   2085
      Left            =   90
      ScaleHeight     =   2085
      ScaleWidth      =   5205
      TabIndex        =   21
      Top             =   4320
      Width           =   5205
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1860
         TabIndex        =   22
         Top             =   30
         Width           =   3255
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1860
         TabIndex        =   24
         Top             =   450
         Width           =   3255
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1860
         TabIndex        =   26
         Top             =   840
         Width           =   3255
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1860
         TabIndex        =   28
         Top             =   1230
         Width           =   3255
      End
      Begin VB.TextBox txtNotedBy2 
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1860
         TabIndex        =   30
         Top             =   1620
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Prepared/ Prepared By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   465
         Left            =   30
         TabIndex        =   23
         Top             =   60
         Width           =   1875
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Checked By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   30
         TabIndex        =   25
         Top             =   570
         Width           =   1725
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   30
         TabIndex        =   27
         Top             =   960
         Width           =   1725
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "1st Noted By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   30
         TabIndex        =   29
         Top             =   1320
         Width           =   1725
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Noted By"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   30
         TabIndex        =   31
         Top             =   1710
         Width           =   1725
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   90
      ScaleHeight     =   1815
      ScaleWidth      =   5205
      TabIndex        =   11
      Top             =   2220
      Width           =   5205
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
         ForeColor       =   &H00701E2A&
         Height          =   375
         Left            =   1890
         TabIndex        =   12
         Top             =   90
         Width           =   3225
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
         ForeColor       =   &H00701E2A&
         Height          =   375
         Left            =   1890
         TabIndex        =   17
         Top             =   900
         Width           =   3225
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1890
         TabIndex        =   15
         Top             =   510
         Width           =   3225
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
         ForeColor       =   &H00701E2A&
         Height          =   375
         Left            =   1890
         TabIndex        =   19
         Top             =   1320
         Width           =   3225
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "General Manager"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   60
         TabIndex        =   13
         Top             =   150
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Manager"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   960
         Width           =   1725
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   570
         Width           =   1725
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Secretary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   30
         TabIndex        =   18
         Top             =   1380
         Width           =   1725
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   90
      ScaleHeight     =   1695
      ScaleWidth      =   7065
      TabIndex        =   1
      Top             =   270
      Width           =   7065
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   7
         Top             =   840
         Width           =   2265
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1890
         MaxLength       =   100
         TabIndex        =   5
         Top             =   450
         Width           =   5055
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   3
         Top             =   60
         Width           =   5055
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1230
         Width           =   2265
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company TIN No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   60
         TabIndex        =   6
         Top             =   870
         Width           =   1785
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   60
         TabIndex        =   4
         Top             =   510
         Width           =   1785
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   60
         TabIndex        =   2
         Top             =   120
         Width           =   1785
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company SSS No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   60
         TabIndex        =   8
         Top             =   1260
         Width           =   1785
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
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
      Left            =   6450
      MouseIcon       =   "Profile.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "Profile.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Close Window"
      Top             =   5580
      Width           =   705
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Update"
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
      Left            =   5760
      MouseIcon       =   "Profile.frx":12D2
      MousePointer    =   99  'Custom
      Picture         =   "Profile.frx":1424
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Save Changes"
      Top             =   5580
      Width           =   705
   End
   Begin VB.Label Label16 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Signatories:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   90
      TabIndex        =   20
      Top             =   4080
      Width           =   3345
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Officers Information:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   90
      TabIndex        =   10
      Top             =   1980
      Width           =   3345
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Information:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   3345
   End
End
Attribute VB_Name = "frmPMISProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSPROFILE                                                         As ADODB.Recordset

Sub StoreMemvars()
    If Not RSPROFILE.EOF And Not RSPROFILE.BOF Then
        txtCompanyName.Text = Null2String(RSPROFILE!CompanyName)
        txtCompanyAddress.Text = Null2String(RSPROFILE!Companyaddress)
        txtCompanyTINNo.Text = Null2String(RSPROFILE!companytinno)
        txtCompanySSSNo.Text = Null2String(RSPROFILE!companysssno)
        txtPreparedBy.Text = Null2String(RSPROFILE!PreparedBy)
        txtCheckedBy.Text = Null2String(RSPROFILE!CheckedBy)
        txtApprovedBy.Text = Null2String(RSPROFILE!ApprovedBy)
        txtNotedBy1.Text = Null2String(RSPROFILE!notedby1)
        txtNotedBy2.Text = Null2String(RSPROFILE!notedby2)
        txtGeneralManager.Text = Null2String(RSPROFILE!GeneralManager)
        txtAccountNo.Text = Null2String(RSPROFILE!ACCOUNTNO)
        txtSBManager.Text = Null2String(RSPROFILE!bankmanager)
        txtCorpSec.Text = Null2String(RSPROFILE!SECRETARY)

        COMPANY_NAME = Null2String(RSPROFILE!CompanyName)
        COMPANY_ADDRESS = Null2String(RSPROFILE!Companyaddress)
        COMPANY_TIN = Null2String(RSPROFILE!companytinno)

        PREPARED_BY = Null2String(RSPROFILE!PreparedBy)
        CHECKED_BY = Null2String(RSPROFILE!CheckedBy)
        APPROVED_BY = Null2String(RSPROFILE!ApprovedBy)

        GENERAL_MANAGER = Null2String(RSPROFILE!GeneralManager)
        ACCOUNT_NO = Null2String(RSPROFILE!ACCOUNTNO)
        BANK_MANAGER = Null2String(RSPROFILE!bankmanager)
        SECRETARY = Null2String(RSPROFILE!SECRETARY)
    End If
End Sub

Sub rsRefresh()
    Set RSPROFILE = New ADODB.Recordset
    RSPROFILE.Open "select * from ALL_Profile WHERE MODULENAME='" & MODULENAME & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ERRORCODE
    If txtCompanyName.Text = "" Then
        MessagePop RecSaveError, "Company must have a Name.", "Invalid Name"
        Exit Sub
    End If
    If txtCompanyAddress.Text = "" Then

        MessagePop RecSaveError, "Company must have a Address.", "Invalid Address", 2500
        Exit Sub
    End If
    If txtCompanyTINNo.Text = "" Then
        If MsgBox("TIN No. Omitted. Proceed Anyway?", vbYesNo + vbQuestion, "Warning") = vbCancel Then
            Exit Sub
        End If
    End If
    If txtCompanySSSNo.Text = "" Then
        If MsgBox("SSS No. Omitted. Proceed Anyway?", vbYesNo + vbQuestion, "Warning") = vbCancel Then
            Exit Sub
        End If
    End If

    SQL_STATEMENT = "UPDATE ALL_PROFILE SET" & _
                   " COMPANYNAME = " & N2Str2Null(txtCompanyName.Text) & "," & _
                   " COMPANYADDRESS = " & N2Str2Null(txtCompanyAddress.Text) & "," & _
                   " COMPANYTINNO = " & N2Str2Null(txtCompanyTINNo.Text) & "," & _
                   " COMPANYSSSNO = " & N2Str2Null(txtCompanySSSNo.Text) & "," & _
                   " PREPAREDBY = " & N2Str2Null(txtPreparedBy.Text) & "," & _
                   " CHECKEDBY = " & N2Str2Null(txtCheckedBy.Text) & "," & _
                   " APPROVEDBY = " & N2Str2Null(txtApprovedBy.Text) & "," & _
                   " NOTEDBY1 = " & N2Str2Null(txtNotedBy1.Text) & "," & _
                   " NOTEDBY2 = " & N2Str2Null(txtNotedBy2.Text) & "," & _
                   " GENERALMANAGER = " & N2Str2Null(txtGeneralManager.Text) & "," & _
                   " ACCOUNTNO = " & N2Str2Null(txtAccountNo.Text) & "," & _
                   " BANKMANAGER = " & N2Str2Null(txtSBManager.Text) & "," & _
                   " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                   " LASTUPDATE = " & N2Str2Null(LOGDATE) & "," & _
                   " SECRETARY = " & N2Str2Null(txtCorpSec.Text) & " WHERE MODULENAME='" & MODULENAME & "'"
    
    gconDMIS.Execute SQL_STATEMENT
    
    NEW_LogAudit "E", "COMPANY PROFILE", SQL_STATEMENT, "", "", "", "", ""
    ShowSuccessFullyUpdated
    Exit Sub

ERRORCODE:
    ShowVBError
    Exit Sub

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
             
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (COMPANY PROFILE)"
            Call frmALL_AuditInquiry.DisplayHistory("", "COMPANY PROFILE", "PRINTING")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    StoreMemvars
    Screen.MousePointer = 0
End Sub


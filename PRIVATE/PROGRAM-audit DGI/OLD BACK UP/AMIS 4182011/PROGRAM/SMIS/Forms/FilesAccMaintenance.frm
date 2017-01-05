VERSION 5.00
Begin VB.Form frmSMIS_FilesAccMaintenance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password File Maintenance"
   ClientHeight    =   3720
   ClientLeft      =   5415
   ClientTop       =   2385
   ClientWidth     =   4200
   ForeColor       =   &H00D8E9EC&
   Icon            =   "FilesAccMaintenance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3720
   ScaleWidth      =   4200
   Begin VB.TextBox txtConfirm 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   22
      PasswordChar    =   "l"
      TabIndex        =   9
      Top             =   2160
      Width           =   3645
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   22
      PasswordChar    =   "l"
      TabIndex        =   8
      Top             =   1530
      Width           =   3645
   End
   Begin VB.TextBox txtCurrentpass 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   22
      PasswordChar    =   "l"
      TabIndex        =   7
      Top             =   870
      Width           =   3645
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   2340
      ScaleHeight     =   885
      ScaleWidth      =   1650
      TabIndex        =   0
      Top             =   2670
      Width           =   1650
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
         Left            =   900
         MouseIcon       =   "FilesAccMaintenance.frx":1472
         MousePointer    =   99  'Custom
         Picture         =   "FilesAccMaintenance.frx":15C4
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Cancel"
         Top             =   60
         Width           =   735
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
         Left            =   180
         MouseIcon       =   "FilesAccMaintenance.frx":1902
         MousePointer    =   99  'Custom
         Picture         =   "FilesAccMaintenance.frx":1A54
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Save Changes"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Type the new password again to confirm"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   3315
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Type your Current Password"
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
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Type a new password"
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
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1260
      Width           =   3375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Change the  password of your account"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   150
      Width           =   3150
   End
End
Attribute VB_Name = "frmSMIS_FilesAccMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    On Error GoTo ErrorCode:

    If LTrim(RTrim(txtCurrentpass)) = "" Then
        ShowIsRequiredMsg " Current Password"
        On Error Resume Next
        txtCurrentpass.SetFocus
        Exit Sub
    End If
    If LTrim(RTrim(txtPassword)) = "" Then
        ShowIsRequiredMsg " New Password"
        On Error Resume Next
        txtPassword.SetFocus
        Exit Sub
    End If
    If LTrim(RTrim(txtConfirm)) = "" Then
        ShowIsRequiredMsg " Confirm Password"
        On Error Resume Next
        txtConfirm.SetFocus
        Exit Sub
    End If

    Dim rspass                                                        As ADODB.Recordset
    Set rspass = gconDMIS.Execute("SELECT PASSWORD FROM SMIS_SALESTEAM WHERE SAECODE=" & N2Str2Null(LOGSAE))
    If Not rspass.EOF Or Not rspass.BOF Then
        If Null2String(rspass!Password) <> txtCurrentpass Then
            MessagePop RecSaveError, "Password Mis Match ", " Password Doesn't Match Your Current Password"
            On Error Resume Next
            txtCurrentpass.SetFocus
            Exit Sub
        End If
    Else
        MsgBox " You Password has not been Configured Yet. " & vbCrLf & " Please Configure Your Password From SALES Department", vbInformation
        Exit Sub
    End If


    If txtConfirm <> txtPassword Then
        MessagePop RecSaveError, "Password Mis Match ", " New Password Doesn't Match with Your Confirm Password"
        On Error Resume Next
        txtConfirm.SetFocus
        Exit Sub
    End If




    gconDMIS.Execute ("Update SMIS_SALESTEAM SET PASSWORD='" & txtPassword & "' WHERE SAECODE=" & LOGSAE)
    ShowSuccessFullyUpdated

    cmdCancel.Value = True

    Exit Sub
ErrorCode:
    ShowVBError


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub


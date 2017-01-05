VERSION 5.00
Begin VB.Form frmHRMSBEGBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Beginning Balances"
   ClientHeight    =   2415
   ClientLeft      =   315
   ClientTop       =   645
   ClientWidth     =   6135
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2205
      Left            =   0
      ScaleHeight     =   2205
      ScaleWidth      =   7965
      TabIndex        =   4
      Top             =   0
      Width           =   7965
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4650
         MaxLength       =   50
         TabIndex        =   12
         Top             =   420
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4650
         MaxLength       =   50
         TabIndex        =   11
         Top             =   780
         Width           =   1185
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4650
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1140
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2970
         MaxLength       =   50
         TabIndex        =   9
         Top             =   60
         Width           =   1185
      End
      Begin VB.TextBox txtMPlaceOfBirth 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1140
         Width           =   1185
      End
      Begin VB.TextBox txtMother 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   2
         Top             =   780
         Width           =   1185
      End
      Begin VB.TextBox txtFPlaceOfBirth 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   1
         Top             =   420
         Width           =   1185
      End
      Begin VB.TextBox txtFather 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   0
         Top             =   60
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   5160
         Picture         =   "BegBalance.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Cancel"
         Top             =   1530
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4470
         Picture         =   "BegBalance.frx":033E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Save Entry"
         Top             =   1530
         Width           =   705
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Pag-Ibig"
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
         Left            =   3000
         TabIndex        =   15
         Top             =   450
         Width           =   2025
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Philhealth"
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
         Left            =   3000
         TabIndex        =   14
         Top             =   810
         Width           =   2595
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "13th Month"
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
         Left            =   3000
         TabIndex        =   13
         Top             =   1170
         Width           =   2025
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SSS"
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
         Left            =   60
         TabIndex        =   8
         Top             =   1170
         Width           =   2025
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Withholding Tax"
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
         Left            =   60
         TabIndex        =   7
         Top             =   810
         Width           =   2595
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Net Taxable"
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
         Left            =   60
         TabIndex        =   6
         Top             =   450
         Width           =   2025
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Range"
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
         Left            =   60
         TabIndex        =   5
         Top             =   90
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmHRMSBEGBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEMPINFO                                As ADODB.Recordset
Dim EMPLIVIL                                 As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo Errorcode:

If Function_Access(LOGID, "Acess_Edit") = False Then Exit Sub

    gconDMIS.Execute "update HRMS_EmpInfo Set" & _
                   " Father = " & N2Str2Null(txtFather.Text) & "," & _
                   " FPlaceofBirth = " & N2Str2Null(txtFPlaceOfBirth.Text) & "," & _
                   " Mother = " & N2Str2Null(txtMother.Text) & "," & _
                   " MPlaceofBirth = " & N2Str2Null(txtMPlaceOfBirth.Text) & _
                   " Where EMPLEVEL = " & EMPLIVIL & " AND Empno = " & EMPLOYEE_NO
    ShowSuccessFullyUpdated
    Unload Me





Exit Sub
Errorcode:
ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 0
    If EMP_TYPE = "EMPLOYEE" Then
        If HEADOREMP = "HEAD" Then
            EMPLIVIL = "'M'"
        Else
            EMPLIVIL = "'E'"
        End If
    End If
    If EMP_TYPE = "CONTRACTUAL" Then EMPLIVIL = "'C'"
    If EMP_TYPE = "ALLOWANCE BASE" Then EMPLIVIL = "'A'"
    rsRefreshStoreMemVars
End Sub

Sub rsRefreshStoreMemVars()
    Set rsEMPINFO = New ADODB.Recordset
    Set rsEMPINFO = gconDMIS.Execute("Select * from HRMS_EmpInfo Where EMPLEVEL = " & EMPLIVIL & " AND Empno = " & EMPLOYEE_NO)
    If Not rsEMPINFO.EOF And Not rsEMPINFO.BOF Then
        txtFather.Text = Null2String(rsEMPINFO!Father)
        txtFPlaceOfBirth.Text = Null2String(rsEMPINFO!FPlaceOfBirth)
        txtMother.Text = Null2String(rsEMPINFO!Mother)
        txtMPlaceOfBirth.Text = Null2String(rsEMPINFO!MPlaceOfBirth)
    End If
End Sub


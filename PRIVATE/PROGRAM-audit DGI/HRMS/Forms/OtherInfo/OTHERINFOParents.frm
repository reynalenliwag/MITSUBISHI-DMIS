VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmOTHERINFOParents 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PARENTS"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5820
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2310
      Left            =   135
      ScaleHeight     =   2250
      ScaleWidth      =   5535
      TabIndex        =   4
      Top             =   105
      Width           =   5595
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
         Height          =   675
         Left            =   4770
         MouseIcon       =   "OTHERINFOParents.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOParents.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancel Entry"
         Top             =   1530
         Width           =   705
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1140
         Width           =   4065
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   2
         Top             =   780
         Width           =   4065
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   1
         Top             =   420
         Width           =   4065
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   0
         Top             =   60
         Width           =   4065
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
         Height          =   675
         Left            =   4080
         MouseIcon       =   "OTHERINFOParents.frx":0490
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOParents.frx":05E2
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Save Entry"
         Top             =   1530
         Width           =   705
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Place of Birth"
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
         TabIndex        =   9
         Top             =   1170
         Width           =   2025
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Mother"
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
         Top             =   810
         Width           =   885
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Place of Birth"
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
         Caption         =   "Father"
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
         Width           =   885
      End
   End
   Begin wizButton.cmd cmdDependents 
      Height          =   2430
      Left            =   90
      TabIndex        =   7
      Top             =   60
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   4286
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "OTHERINFOParents.frx":0932
   End
End
Attribute VB_Name = "frmOTHERINFOParents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSEMPINFO                                                         As ADODB.Recordset
Dim EMPLIVIL                                                          As String

Sub rsRefreshStoreMemVars()
    Set RSEMPINFO = New ADODB.Recordset
    Set RSEMPINFO = gconDMIS.Execute("Select * from HRMS_EmpInfo Where EMPLEVEL = " & EMPLIVIL & " AND Empno = " & EMPLOYEE_NO)
    If Not RSEMPINFO.EOF And Not RSEMPINFO.BOF Then
        txtFather.Text = Null2String(RSEMPINFO!FATHER)
        txtFPlaceOfBirth.Text = Null2String(RSEMPINFO!FPlaceOfBirth)
        txtMother.Text = Null2String(RSEMPINFO!MOTHER)
        txtMPlaceOfBirth.Text = Null2String(RSEMPINFO!MPlaceOfBirth)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200711:58
Private Sub cmdSave_Click()
    On Error GoTo Errorcode:

    gconDMIS.Execute "update HRMS_EmpInfo Set" & _
                   " Father = " & N2Str2Null(txtFather.Text) & "," & _
                   " FPlaceofBirth = " & N2Str2Null(txtFPlaceOfBirth.Text) & "," & _
                   " Mother = " & N2Str2Null(txtMother.Text) & "," & _
                   " MPlaceofBirth = " & N2Str2Null(txtMPlaceOfBirth.Text) & _
                   " Where EMPLEVEL = " & EMPLIVIL & " AND Empno = " & EMPLOYEE_NO

    Call LogAudit("E", "UPDATE EMPLOYEE PARENTS INFORMATION", EMPLOYEE_NO)
    Call ShowSuccessFullyUpdated
    Unload Me

    Exit Sub

Errorcode:
    Call ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 0
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
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


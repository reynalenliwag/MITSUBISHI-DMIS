VERSION 5.00
Begin VB.Form frmPMISSignatories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Configuration"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Signatories.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   6780
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   5175
      ScaleHeight     =   885
      ScaleWidth      =   1980
      TabIndex        =   20
      Top             =   4425
      Width           =   1980
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
         Left            =   765
         MouseIcon       =   "Signatories.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Signatories.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   30
         MouseIcon       =   "Signatories.frx":079A
         MousePointer    =   99  'Custom
         Picture         =   "Signatories.frx":08EC
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "SYSTEM SIGNATORIES"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   120
      TabIndex        =   9
      Top             =   1950
      Width           =   6525
      Begin VB.TextBox txtPREPARED_BY 
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
         ForeColor       =   &H00000040&
         Height          =   345
         Left            =   2370
         TabIndex        =   14
         Text            =   "Text1"
         ToolTipText     =   "Input prepared by name."
         Top             =   1110
         Width           =   4065
      End
      Begin VB.TextBox txtPARTS_MANAGER 
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
         ForeColor       =   &H00000040&
         Height          =   345
         Left            =   2370
         TabIndex        =   13
         Text            =   "Text1"
         ToolTipText     =   "Input approved by name."
         Top             =   1500
         Width           =   4065
      End
      Begin VB.TextBox txtGENERAL_MANAGER 
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
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   2370
         TabIndex        =   12
         Text            =   "Text1"
         ToolTipText     =   "Type the name of the general manager."
         Top             =   1890
         Width           =   4065
      End
      Begin VB.TextBox txtCOUNTER_OFFICER 
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
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   2370
         TabIndex        =   11
         Text            =   "Text1"
         ToolTipText     =   "Type the name of the general manager."
         Top             =   690
         Width           =   4065
      End
      Begin VB.TextBox txtWAREHOUSE_OFFICER 
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
         ForeColor       =   &H00000040&
         Height          =   345
         Left            =   2370
         TabIndex        =   10
         Text            =   "Text1"
         ToolTipText     =   "Input issued by name."
         Top             =   300
         Width           =   4065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "REPORTS PREPARED BY"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1140
         Width           =   2205
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "PARTS MANAGER"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1530
         Width           =   2115
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "GENERAL MANAGER"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   2235
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "COUNTER OFFICER"
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
         Height          =   435
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   2265
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "WAREHOUSE OFFICER"
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
         Height          =   435
         Left            =   120
         TabIndex        =   15
         Top             =   330
         Width           =   2265
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DEALER SETTINGS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6525
      Begin VB.TextBox txtDEALERCODE 
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
         ForeColor       =   &H00000040&
         Height          =   345
         Left            =   2370
         TabIndex        =   5
         Text            =   "Text1"
         ToolTipText     =   "Input prepared by name."
         Top             =   240
         Width           =   1305
      End
      Begin VB.TextBox txtCOMPANY_NAME 
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
         ForeColor       =   &H00000040&
         Height          =   345
         Left            =   2370
         TabIndex        =   4
         Text            =   "Text1"
         ToolTipText     =   "Input issued by name."
         Top             =   600
         Width           =   4065
      End
      Begin VB.TextBox txtBUSINESS_ADDRESS 
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
         ForeColor       =   &H00000040&
         Height          =   765
         Left            =   2370
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "Signatories.frx":0C3C
         ToolTipText     =   "Input approved by name."
         Top             =   960
         Width           =   4065
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "DEALER CODE"
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
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   1725
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "COMPANY NAME"
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
         Left            =   120
         TabIndex        =   7
         Top             =   630
         Width           =   1725
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "BUSINESS ADDRESS"
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
         Height          =   675
         Left            =   120
         TabIndex        =   6
         Top             =   990
         Width           =   2085
      End
      Begin VB.Label labPrev 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         Height          =   345
         Left            =   150
         TabIndex        =   2
         Top             =   270
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label labid 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   15
      End
   End
End
Attribute VB_Name = "frmPMISSignatories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSignatories        As ADODB.Recordset
Dim AddorEdit            As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode

    txtPREPARED_BY.Text = N2Str2Null(txtPREPARED_BY.Text)
    txtWAREHOUSE_OFFICER.Text = N2Str2Null(txtWAREHOUSE_OFFICER.Text)
    txtPARTS_MANAGER.Text = N2Str2Null(txtPARTS_MANAGER.Text)
    txtGENERAL_MANAGER.Text = N2Str2Null(txtGENERAL_MANAGER.Text)
    gconDMIS.Execute "update Signatories set" & _
                   " preparedby = " & txtPREPARED_BY.Text & "," & _
                   " IssuedBy = " & txtWAREHOUSE_OFFICER.Text & "," & _
                   " approvedby = " & txtPARTS_MANAGER.Text & "," & _
                   " generalmanager = " & txtGENERAL_MANAGER.Text & _
                   " where id = " & labid.Caption
    rsRefresh
    On Error Resume Next
    rsSignatories.Find "id = " & labid.Caption
    cmdCancel.Value = True
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    rsRefresh
    SetFormSettings Me
    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub initMemvars()
    txtPREPARED_BY.Text = ""
    txtWAREHOUSE_OFFICER.Text = ""
    txtPARTS_MANAGER.Text = ""
    txtGENERAL_MANAGER.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsSignatories.EOF And Not rsSignatories.BOF Then
        labid.Caption = rsSignatories!ID
        txtPREPARED_BY.Text = Null2String(rsSignatories!preparedby)
        txtWAREHOUSE_OFFICER.Text = Null2String(rsSignatories!issuedby)
        txtPARTS_MANAGER.Text = Null2String(rsSignatories!approvedby)
        txtGENERAL_MANAGER.Text = Null2String(rsSignatories!generalmanager)
    Else
        ShowNoRecord
    End If
End Sub

Sub rsRefresh()
    Set rsSignatories = New ADODB.Recordset
    rsSignatories.Open "select * from Signatories", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISSignatories = Nothing
    UnloadForm Me
End Sub
